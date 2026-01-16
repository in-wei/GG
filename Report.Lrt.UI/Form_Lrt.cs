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
using System.IO;
using UsuallyCommond.MyEnum;
using UsuallyCommond;
//using Report.Lrt.ViewModel;
using Excel = Microsoft.Office.Interop.Excel;
using GlobalCommond;
using GlobalCommond.ViewModel;

namespace Report.Lrt.UI
{
    public partial class Form_Lrt : Form
    {
        #region 變數定義
        private MyCommonds MyCommond;
        private LrtMain MainServer;
        private string ThisReceive;

        List<Timers> MyTimers;
        List<ProgressBars> MyProgressBars;
        List<CheckBoxs> MyCheckBoxs;

        private bool ConfigLoading = false;
        public bool Test_Button = false;

        private bool CheckPaoCsv = false;
        public bool LoadingStatu = false;

        List<PaoTicketSaleList> Form_PaoText_UI_Status;
        List<PaoTicketSaleList> Form_PaoText_Data;

        Timer Text_timer;

        #endregion

        public Form_Lrt()
        {
            InitializeComponent();
            ThisReceive = "Form_" + this.Text;
            MainServer = new LrtMain();
            MainServer.globalCommond = new GlobalCommonds(true, true, true);
            //MainServer.globalCommond = new GlobalCommonds("LRT", true, true, true);
            MainServer.SetMyCommand();
            //MyCommond = MainServer.globalCommond.MyCommond;
            MyCommond = new MyCommonds();
            MyCommond.WriteLog(ThisReceive, "==程式啟動==");
            string ProgramFullPath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName.ToString();
            string ProgramLastWriteTime = File.GetLastWriteTime(ProgramFullPath).ToString();
            MyCommond.WriteLog(ThisReceive, $"This Program Last Write Time： {ProgramLastWriteTime}");

            if (MainServer.globalCommond.UsingSystem.Equals("lrt", StringComparison.CurrentCultureIgnoreCase) && MainServer.globalCommond.OperationCode_String != null)
            {
                RegisterEvents();
                Other_Init();
            }
            else
            {
                MessageBox.Show($@"設定檔中未發現輕軌設定，請檢查下列檔案{MyCommond.Path_Program}{MyCommond.Path_Config}\{MyUseFile.Config_Setting}"
                        , "錯誤"
                        , MessageBoxButtons.OK
                        , MessageBoxIcon.Question);
                GC.Collect();
                MyCommond.WriteLog(ThisReceive, "程式關閉");
                System.Environment.Exit(0);
            }
        }

        private void RegisterEvents()
        {
            MyCommond.WriteLog(ThisReceive, "掛聽各物件");

            //Form
            this.Load += Form_Load;
            this.MinimumSizeChanged += Form_MinimumSizeChanged;
            this.FormClosing += Form_Close;
            this.SizeChanged += Form_SizeChanged;

            // Menu
            this.MenuItem_Close.Click += MenuItem_Close_Click;

            //UI

            #region Txet

            #region Group Pao+

            #region Pao1

            textBox_OneDay_1.KeyPress += textBox_OneDay_1_Changed;
            textBox_OneDay_E_1.KeyPress += textBox_OneDay_E_1_Changed;
            textBox_Ticket_10_1.KeyPress += textBox_Ticket_10_1_Changed;
            textBox_Ticket_12_1.KeyPress += textBox_Ticket_12_1_Changed;
            textBox_Ticket_15_1.KeyPress += textBox_Ticket_15_1_Changed;
            textBox_Ticket_20_1.KeyPress += textBox_Ticket_20_1_Changed;
            textBox_Ticket_25_1.KeyPress += textBox_Ticket_25_1_Changed;
            textBox_Ticket_30_1.KeyPress += textBox_Ticket_30_1_Changed;
            textBox_Ticket_50_1.KeyPress += textBox_Ticket_50_1_Changed;
            textBox_OpenTransaction_1.KeyPress += textBox_OpenTransaction_1_Changed;
            textBox_Tpass_1.KeyPress += textBox_Tpass_1_Changed;
            textBox_Group_Count_1.KeyPress += textBox_Group_Count_1_Changed;
            textBox_Group_People_1.KeyPress += textBox_Group_People_1_Changed;
            textBox_Group_Amount_1.KeyPress += textBox_Group_Amount_1_Changed;
            textBox_Lanwair1_1.KeyPress += textBox_Lanwair1_1_Changed;
            textBox_Lanwair2_1.KeyPress += textBox_Lanwair2_1_Changed;
            textBox_VavmTake_1.KeyPress += textBox_VavmTake_1_Changed;
            textBox_Repair_1.KeyPress += textBox_Repair_1_Changed;

            #endregion

            #region Pao2

            textBox_OneDay_2.KeyPress += textBox_OneDay_2_Changed;
            textBox_OneDay_E_2.KeyPress += textBox_OneDay_E_2_Changed;
            textBox_Ticket_10_2.KeyPress += textBox_Ticket_10_2_Changed;
            textBox_Ticket_12_2.KeyPress += textBox_Ticket_12_2_Changed;
            textBox_Ticket_15_2.KeyPress += textBox_Ticket_15_2_Changed;
            textBox_Ticket_20_2.KeyPress += textBox_Ticket_20_2_Changed;
            textBox_Ticket_25_2.KeyPress += textBox_Ticket_25_2_Changed;
            textBox_Ticket_30_2.KeyPress += textBox_Ticket_30_2_Changed;
            textBox_Ticket_50_2.KeyPress += textBox_Ticket_50_2_Changed;
            textBox_OpenTransaction_2.KeyPress += textBox_OpenTransaction_2_Changed;
            textBox_Tpass_2.KeyPress += textBox_Tpass_2_Changed;
            textBox_Group_Count_2.KeyPress += textBox_Group_Count_2_Changed;
            textBox_Group_People_2.KeyPress += textBox_Group_People_2_Changed;
            textBox_Group_Amount_2.KeyPress += textBox_Group_Amount_2_Changed;
            textBox_Lanwair1_2.KeyPress += textBox_Lanwair1_2_Changed;
            textBox_Lanwair2_2.KeyPress += textBox_Lanwair2_2_Changed;
            textBox_VavmTake_2.KeyPress += textBox_VavmTake_2_Changed;
            textBox_Repair_2.KeyPress += textBox_Repair_2_Changed;

            #endregion

            #endregion

            #region Group Pao-

            #region Pao1

            textBox_ConcessionRefund_1.KeyPress += textBox_ConcessionRefund_1_Changed;
            textBox_RefundNotice_1.KeyPress += textBox_RefundNotice_1_Changed;
            textBox_VavmRefund_1.KeyPress += textBox_VavmRefund_1_Changed;
            textBox_PvRefund_1.KeyPress += textBox_PvRefund_1_Changed;
            textBox_CreditCardRefund_1.KeyPress += textBox_CreditCardRefund_1_Changed;
            textBox_QrCodeRefund_1.KeyPress += textBox_QrCodeRefund_1_Changed;
            textBox_SaleCardRefund_1.KeyPress += textBox_SaleCardRefund_1_Changed;
            textBox_GroupRefund_1.KeyPress += textBox_GroupRefund_1_Changed;

            #endregion

            #region Pao2

            textBox_ConcessionRefund_2.KeyPress += textBox_ConcessionRefund_2_Changed;
            textBox_RefundNotice_2.KeyPress += textBox_RefundNotice_2_Changed;
            textBox_VavmRefund_2.KeyPress += textBox_VavmRefund_2_Changed;
            textBox_PvRefund_2.KeyPress += textBox_PvRefund_2_Changed;
            textBox_CreditCardRefund_2.KeyPress += textBox_CreditCardRefund_2_Changed;
            textBox_QrCodeRefund_2.KeyPress += textBox_QrCodeRefund_2_Changed;
            textBox_SaleCardRefund_2.KeyPress += textBox_SaleCardRefund_2_Changed;
            textBox_GroupRefund_2.KeyPress += textBox_GroupRefund_2_Changed;

            #endregion

            #endregion

            #region Group OneDay

            textBox_ExchangeOneDay_1.KeyPress += textBox_ExchangeOneDay_1_Changed;
            textBox_ExchangeOneDay_2.KeyPress += textBox_ExchangeOneDay_2_Changed;
            textBox_KKDay.KeyPress += textBox_KKDay_Changed;
            textBox_Ntpc.KeyPress += textBox_Ntpc_Changed;
            textBox_SplitTicker.KeyPress += textBox_SplitTicker_Changed;

            #endregion

            #region Group Exchange

            #region Pao1

            textBox_ExchangeTicket_20_1.KeyPress += textBox_ExchangeTicket_20_1_Changed;
            textBox_ExchangeTicket_25_1.KeyPress += textBox_ExchangeTicket_25_1_Changed;
            textBox_ExchangeTicket_30_1.KeyPress += textBox_ExchangeTicket_30_1_Changed;

            #endregion

            #region Pao2

            textBox_ExchangeTicket_20_2.KeyPress += textBox_ExchangeTicket_20_2_Changed;
            textBox_ExchangeTicket_25_2.KeyPress += textBox_ExchangeTicket_25_2_Changed;
            textBox_ExchangeTicket_30_2.KeyPress += textBox_ExchangeTicket_30_2_Changed;

            #endregion

            #endregion

            #region Other

            textBox0.KeyPress += TextBox0_Changed;

            #endregion

            #endregion

            #region CheckBox

            #region Group1

            checkBox1.CheckedChanged += checkBox1_CheckedChanged;
            checkBox2.CheckedChanged += checkBox2_CheckedChanged;
            checkBox3.CheckedChanged += checkBox3_CheckedChanged;
            checkBox4.CheckedChanged += checkBox4_CheckedChanged;

            #endregion

            #region Group2

            checkBox_G_1_1.CheckedChanged += checkBox_G_1_1_CheckedChanged;
            checkBox_G_1_2.CheckedChanged += checkBox_G_1_2_CheckedChanged;
            checkBox_G_1_3.CheckedChanged += checkBox_G_1_3_CheckedChanged;
            checkBox_G_1_4.CheckedChanged += checkBox_G_1_4_CheckedChanged;
            checkBox_G_1_5.CheckedChanged += checkBox_G_1_5_CheckedChanged;
            checkBox_G_1_6.CheckedChanged += checkBox_G_1_6_CheckedChanged;
            checkBox_G_1_7.CheckedChanged += checkBox_G_1_7_CheckedChanged;
            checkBox_G_1_8.CheckedChanged += checkBox_G_1_8_CheckedChanged;

            #endregion

            #region Group3

            checkBox_G_2_1.CheckedChanged += checkBox_G_2_1_CheckedChanged;
            checkBox_G_2_2.CheckedChanged += checkBox_G_2_2_CheckedChanged;
            checkBox_G_2_3.CheckedChanged += checkBox_G_2_3_CheckedChanged;
            checkBox_G_2_4.CheckedChanged += checkBox_G_2_4_CheckedChanged;
            checkBox_G_2_5.CheckedChanged += checkBox_G_2_5_CheckedChanged;
            checkBox_G_2_6.CheckedChanged += checkBox_G_2_6_CheckedChanged;
            checkBox_G_2_7.CheckedChanged += checkBox_G_2_7_CheckedChanged;
            checkBox_G_2_8.CheckedChanged += checkBox_G_2_8_CheckedChanged;
            checkBox_G_2_9.CheckedChanged += checkBox_G_2_9_CheckedChanged;
            checkBox_G_2_10.CheckedChanged += checkBox_G_2_10_CheckedChanged;
            checkBox_G_2_11.CheckedChanged += checkBox_G_2_11_CheckedChanged;
            checkBox_G_2_12.CheckedChanged += checkBox_G_2_12_CheckedChanged;

            #endregion

            #endregion

            #region DateTimePicker

            dateTimePicker1.ValueChanged += dateTimePicker1_ValueChanged;
            dateTimePicker2.ValueChanged += dateTimePicker2_ValueChanged;
            dateTimePicker3.ValueChanged += dateTimePicker3_ValueChanged;
            dateTimePicker4.ValueChanged += dateTimePicker4_ValueChanged;
            dateTimePicker5.ValueChanged += dateTimePicker5_ValueChanged;
            dateTimePicker6.ValueChanged += dateTimePicker6_ValueChanged;

            #endregion

            #region Button

            #region Group1

            button1.Click += Button1_Click;
            button2.Click += Button2_Click;
            button3.Click += Button3_Click;
            button4.Click += Button4_Click;
            button5.Click += Button5_Click;
            button6.Click += Button6_Click;
            button7.Click += Button7_Click;
            button8.Click += Button8_Click;

            #endregion

            #region Group2

            button9.Click += Button9_Click;
            button10.Click += Button10_Click;
            button11.Click += Button11_Click;
            button12.Click += Button12_Click;
            button13.Click += Button13_Click;
            button14.Click += Button14_Click;
            button15.Click += Button15_Click;
            button16.Click += Button16_Click;
            button17.Click += Button17_Click;
            button18.Click += Button18_Click;
            button19.Click += Button19_Click;
            button20.Click += Button20_Click;

            #endregion

            #region Group3

            button_G_1_All.Click += Button_G_1_All_Click;
            button_G_1_Cancel.Click += Button_G_1_Cancel_Click;
            button_G_1_Start.Click += Button_G_1_Start_Click;

            #endregion

            #region Group4

            button_G_2_All.Click += Button_G_2_All_Click;
            button_G_2_Cancel.Click += Button_G_2_Cancel_Click;
            button_G_2_Start.Click += Button_G_2_Start_Click;

            #endregion

            #region Other

            button_CreditCard.Click += Button_CreditCard_Click;
            button_QRCode.Click += Button_QRCode_Click;
            button_Test.Click += Button_Test_Click;

            #endregion

            #endregion



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

        #region Form - Event

        public void Form_Load(object sender, EventArgs e)
        {
            ConfigLoading = true;
            MyCommond.WriteLog(ThisReceive, "畫面載入");
            try
            {
                Stock_Set();
                Progerss_Set();
                CheckBox_Set();
                Timer_Set();
                Form_UI_Set();
                SetTipUI();
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR! \n{MyCommond.ExportErrorMessageToLog(ex).Message}");
            }
            ConfigLoading = false;
        }

        private void Form_MinimumSizeChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, "縮到最小");

        }

        public void Form_Close(object sender, FormClosingEventArgs e)
        {
            if (!true)
            {
                this.ShowInTaskbar = false;
                this.WindowState = FormWindowState.Minimized;
                this.Hide();
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
            //MyCommond.WriteLog(ThisReceive, "改變視窗大小");

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
            toolTip1.SetToolTip(groupBox3, "單選");
            toolTip1.SetToolTip(groupBox4, "單選");
        }

        #endregion

        #region 載入

        private void Stock_Set()
        {
            MyCommond.WriteLog(ThisReceive, "初始化變數");
            MyProgressBars = new List<ProgressBars>();
            MyCheckBoxs = new List<CheckBoxs>();
            MyTimers = new List<Timers>();
        }

        private void Progerss_Set()
        {
            MyCommond.WriteLog(ThisReceive, "初始化進度條List");
            for (int i = 1; i <= 20; i++)
            {
                ProgressBar progressBar = new ProgressBar();
                switch (i)
                {
                    case (int)LrtTxnClientSwitch.Day_Volume:                     progressBar = progressBar1;  break;
                    case (int)LrtTxnClientSwitch.Day_Amount:                     progressBar = progressBar2;  break;
                    case (int)LrtTxnClientSwitch.Day_EachStation_EachTime:       progressBar = progressBar3;  break;
                    case (int)LrtTxnClientSwitch.Day_ElectronicTicket:           progressBar = progressBar4;  break;
                    case (int)LrtTxnClientSwitch.Day_AllRideList:                progressBar = progressBar5;  break;
                    case (int)LrtTxnClientSwitch.Day_TrafficAmount:              progressBar = progressBar6;  break;
                    case (int)LrtTxnClientSwitch.Day_OriginDestination:          progressBar = progressBar7;  break;
                    case (int)LrtTxnClientSwitch.Day_EquipAmount:                progressBar = progressBar8;  break;
                    case (int)LrtTxnClientSwitch.Month_AllRideList:              progressBar = progressBar9;  break;
                    case (int)LrtTxnClientSwitch.Month_OriginDestination:        progressBar = progressBar10; break;
                    case (int)LrtTxnClientSwitch.Month_ElectronicTicket_Station: progressBar = progressBar11; break;
                    case (int)LrtTxnClientSwitch.Month_ElectronicTicket_Day:     progressBar = progressBar12; break;
                    case (int)LrtTxnClientSwitch.Month_OwnTicketVolume_Station:  progressBar = progressBar13; break;
                    case (int)LrtTxnClientSwitch.Month_OwnTicketVolume_Day:      progressBar = progressBar14; break;
                    case (int)LrtTxnClientSwitch.Month_TrafficAmount_Station:    progressBar = progressBar15; break;
                    case (int)LrtTxnClientSwitch.Month_TrafficAmount_Day:        progressBar = progressBar16; break;
                    case (int)LrtTxnClientSwitch.Month_EquipAmount_Station:      progressBar = progressBar17; break;
                    case (int)LrtTxnClientSwitch.Month_EquipAmount_Day:          progressBar = progressBar18; break;
                    case (int)LrtTxnClientSwitch.Month_Volume:                   progressBar = progressBar19; break;
                    case (int)LrtTxnClientSwitch.Month_Amount:                   progressBar = progressBar20; break;
                    default: continue;
                }

                MyProgressBars.Add(new ProgressBars()
                { Methode = (LrtTxnClientSwitch)i, Item = progressBar, });
            }
        }

        private void CheckBox_Set()
        {
            MyCommond.WriteLog(ThisReceive, "初始化核選方塊List");
            for (int i = 1; i <= 20; i++)
            {
                CheckBox checkBox = new CheckBox();
                switch (i)
                {
                    case (int)LrtTxnClientSwitch.Day_Volume:                     checkBox = checkBox_G_1_1; break;
                    case (int)LrtTxnClientSwitch.Day_Amount:                     checkBox = checkBox_G_1_2; break;
                    case (int)LrtTxnClientSwitch.Day_EachStation_EachTime:       checkBox = checkBox_G_1_3; break;
                    case (int)LrtTxnClientSwitch.Day_ElectronicTicket:           checkBox = checkBox_G_1_4; break;
                    case (int)LrtTxnClientSwitch.Day_AllRideList:                checkBox = checkBox_G_1_5; break;
                    case (int)LrtTxnClientSwitch.Day_TrafficAmount:              checkBox = checkBox_G_1_6; break;
                    case (int)LrtTxnClientSwitch.Day_OriginDestination:          checkBox = checkBox_G_1_7; break;
                    case (int)LrtTxnClientSwitch.Day_EquipAmount:                checkBox = checkBox_G_1_8; break;
                    case (int)LrtTxnClientSwitch.Month_AllRideList:              checkBox = checkBox_G_2_1; break;
                    case (int)LrtTxnClientSwitch.Month_OriginDestination:        checkBox = checkBox_G_2_2; break;
                    case (int)LrtTxnClientSwitch.Month_ElectronicTicket_Station: checkBox = checkBox_G_2_3; break;
                    case (int)LrtTxnClientSwitch.Month_ElectronicTicket_Day:     checkBox = checkBox_G_2_4; break;
                    case (int)LrtTxnClientSwitch.Month_OwnTicketVolume_Station:  checkBox = checkBox_G_2_5; break;
                    case (int)LrtTxnClientSwitch.Month_OwnTicketVolume_Day:      checkBox = checkBox_G_2_6; break;
                    case (int)LrtTxnClientSwitch.Month_TrafficAmount_Station:    checkBox = checkBox_G_2_7; break;
                    case (int)LrtTxnClientSwitch.Month_TrafficAmount_Day:        checkBox = checkBox_G_2_8; break;
                    case (int)LrtTxnClientSwitch.Month_EquipAmount_Station:      checkBox = checkBox_G_2_9; break;
                    case (int)LrtTxnClientSwitch.Month_EquipAmount_Day:          checkBox = checkBox_G_2_10; break;
                    case (int)LrtTxnClientSwitch.Month_Volume:                   checkBox = checkBox_G_2_11; break;
                    case (int)LrtTxnClientSwitch.Month_Amount:                   checkBox = checkBox_G_2_12; break;
                    default: continue;
                }

                MyCheckBoxs.Add(new CheckBoxs()
                { Methode = (LrtTxnClientSwitch)i, Status = checkBox.Checked, Item = checkBox });
            }


        }

        private void Timer_Set()
        {
            MyCommond.WriteLog(ThisReceive, "初始化時間振盪器List");

            if (!true)
            {
                Text_timer = new Timer();
                Text_timer.Interval = 1000;
                Text_timer.Tick += TextTime;
                Text_timer.Start();
            }
            else
            {
                for (int i = 1; i <= 20; i++)
                {
                    Timer timer = new Timer();
                    timer.Interval = 1000;

                    switch (i)
                    {
                        case (int)LrtTxnClientSwitch.Day_Volume:                     timer.Tick += Timer1;  break;
                        case (int)LrtTxnClientSwitch.Day_Amount:                     timer.Tick += Timer2;  break;
                        case (int)LrtTxnClientSwitch.Day_EachStation_EachTime:       timer.Tick += Timer3;  break;
                        case (int)LrtTxnClientSwitch.Day_ElectronicTicket:           timer.Tick += Timer4;  break;
                        case (int)LrtTxnClientSwitch.Day_AllRideList:                timer.Tick += Timer5;  break;
                        case (int)LrtTxnClientSwitch.Day_TrafficAmount:              timer.Tick += Timer6;  break;
                        case (int)LrtTxnClientSwitch.Day_OriginDestination:          timer.Tick += Timer7;  break;
                        case (int)LrtTxnClientSwitch.Day_EquipAmount:                timer.Tick += Timer8;  break;
                        case (int)LrtTxnClientSwitch.Month_AllRideList:              timer.Tick += Timer9;  break;
                        case (int)LrtTxnClientSwitch.Month_OriginDestination:        timer.Tick += Timer10; break;
                        case (int)LrtTxnClientSwitch.Month_ElectronicTicket_Station: timer.Tick += Timer11; break;
                        case (int)LrtTxnClientSwitch.Month_ElectronicTicket_Day:     timer.Tick += Timer12; break;
                        case (int)LrtTxnClientSwitch.Month_OwnTicketVolume_Station:  timer.Tick += Timer13; break;
                        case (int)LrtTxnClientSwitch.Month_OwnTicketVolume_Day:      timer.Tick += Timer14; break;
                        case (int)LrtTxnClientSwitch.Month_TrafficAmount_Station:    timer.Tick += Timer15; break;
                        case (int)LrtTxnClientSwitch.Month_TrafficAmount_Day:        timer.Tick += Timer16; break;
                        case (int)LrtTxnClientSwitch.Month_EquipAmount_Station:      timer.Tick += Timer17; break;
                        case (int)LrtTxnClientSwitch.Month_EquipAmount_Day:          timer.Tick += Timer18; break;
                        case (int)LrtTxnClientSwitch.Month_Volume:                   timer.Tick += Timer19; break;
                        case (int)LrtTxnClientSwitch.Month_Amount:                   timer.Tick += Timer20; break;
                        default: continue;
                    }
                    timer.Start();
                    MyTimers.Add(new Timers
                    { Methode = (LrtTxnClientSwitch)i, Status = StatusEnum.Stop, Item = timer });
                }
            }
        }

        public void Form_UI_Set()
        {
            MyCommond.WriteLog(ThisReceive, "UI設定");
            dateTimePicker2.Value = DateTime.Now.AddDays(-1);

            dateTimePicker3.Value = DateTime.Now.AddDays(-1);
            dateTimePicker4.Value = DateTime.Now.AddDays(-1);

            DateTime mDt = Convert.ToDateTime(DateTime.Now.ToString("yyyy/MM/01 05:00:00"));
            dateTimePicker5.Value = mDt.AddMonths(-1);
            dateTimePicker6.Value = mDt.AddDays(-1);

            System.Threading.Thread thread = new System.Threading.Thread(() =>
            {
                int i = 0;
                while (i > 0)
                {
                    try
                    {
                        i = MainServer.globalCommond.OperationCode.Count();
                    }
                    catch (Exception ex)
                    {
                        i = 0;
                    }
                    System.Threading.Thread.Sleep(1000);
                }

                MyCommond.InvokeIfRequired(this, () => this.Text += "_" + MainServer.globalCommond.OperationCode);


            });
            thread.IsBackground = true;
            thread.Start();

            LoadItem();
            // All_Item_Status(true); // 測試
        }

        private void LoadItem()
        {
            MyCommond.WriteLog(ThisReceive, $"各物件開放");
            System.Threading.Thread threadLoading = new System.Threading.Thread(() =>
            {
                //Form_PaoText_UI_Status = new List<PaoTicketSaleList>();
                bool TF = true;

                //MainServer.subLoading.Loading_Normal();
                var ssss = MainServer.MyCommond.ExportLanguage;
                //MainServer.subLoading.Loading_Operate();
                //MainServer.subLoading.Loading_Sql();
                MainServer.globalCommond.Stock_Lrt_PaoStation = MainServer.globalCommond.Stock_Lrt_Station.FindAll(x => x.StationPao == "1").ToList();
                //Form_PaoText_UI_Status = MainServer.subLoading.PaoStation();
                Form_PaoText_UI_Status = MainServer.globalCommond.PaoStation();

                this.MyCommond.InvokeIfRequired(this, () =>
                {
                    checkBox2.Checked = MainServer.globalCommond.Excel_Applaction_Show;
                    checkBox3.Checked = MainServer.globalCommond.Auto_Pao_Excel;
                    checkBox4.Checked = MainServer.globalCommond.Install_MobilePay;

                    if (MainServer.globalCommond.Stock_Lrt_PaoStation.Count > 0)
                    {
                        var station = MainServer.globalCommond.Stock_Lrt_PaoStation[0].CodeName;
                        label_pao_1_1.Text = station;
                        label_pao_1_2.Text = station;
                        label_pao_1_3.Text = station;
                        label_pao_1_4.Text = station;
                    }

                    if (MainServer.globalCommond.Stock_Lrt_PaoStation.Count > 1)
                    {
                        var station = MainServer.globalCommond.Stock_Lrt_PaoStation[1].CodeName;
                        label_pao_2_1.Text = station;
                        label_pao_2_2.Text = station;
                        label_pao_2_3.Text = station;
                        label_pao_2_4.Text = station;
                    }
                });

                if (MainServer.globalCommond.Auto_Pao_Excel)
                {
                    MyCommond.WriteLog(ThisReceive, $"準備自動載入班結表");
                    GetPaoSaleList();
                    Data2Form(Form_PaoText_Data);
                }

                Pao_Status(TF);

                if (Form_PaoText_UI_Status.Count > 0)
                {
                    //Pao_Status(1, TF);
                    //if (MainLrt.Stock_Lrt_Station.FindAll(x => x.StationPao == "1").Count() > 1) Pao_Status(2, TF);

                    Group1_Status(TF);
                    Group2_Status(TF);
                    Group3_Status(TF);
                    Button_Status(TF);
                    Other_Status(TF);
                }
            });
            threadLoading.IsBackground = true;
            threadLoading.Start();
        }

        #region PAO

        private bool NeedExport = false;

        private void GetPaoSaleList()
        {
            MyCommond.WriteLog(ThisReceive, $"自動載入班結表初始化");
            List<EachSellReport> PaoSaleList = new List<EachSellReport>(); //存放Pao所有的銷售紀錄
            List<GroupTicketList> GroupList = new List<GroupTicketList>(); //存放所有團體票資料
            Form_PaoText_Data = new List<PaoTicketSaleList>();        //初始化給UI使用的List

            string PaoData_FullPath = $"{MainServer.Path_PaoData}{dateTimePicker2.Value.ToString("yyyyMM")}_{MainServer.FileName_PaoCsv}_{MainServer.globalCommond.OperationCode_String}.csv";
            MyCommond.WriteLog(ThisReceive, $"班結表資料庫路徑:{PaoData_FullPath}");

            if (MyCommond.CheckFile(PaoData_FullPath))
            {
                var data = MyCommond.ToViewModel<PaoTicketSaleList>(ThisReceive, PaoData_FullPath, new PaoTicketSaleList().ExportTitle_zhTW(), 1);
                MyCommond.WriteLog(ThisReceive, $"載入資料筆數{data.Count}");
                Form_PaoText_Data = data.FindAll(x => x.Operation == dateTimePicker2.Value.ToString("yyyyMMdd")).ToList();

                if (Form_PaoText_Data.Count > 0 && Form_PaoText_Data != null) { goto EndSub; }
                MyCommond.WriteLog(ThisReceive, $"未發現本次營運日資料");
            }

            Excel.Application xlApp = new Excel.Application();
            xlApp.Visible = MainServer.globalCommond.Excel_Applaction_Show;

            var HavePaoStation = MainServer.globalCommond.Stock_Lrt_Station.FindAll(x => x.StationPao == "1").ToList();
            MyCommond.WriteLog(ThisReceive, $"當前有詢問處車站數量:{HavePaoStation.Count}");

            foreach (var item in HavePaoStation)
            {
                MyCommond.WriteLog(ThisReceive, $"載入{item.StationName}班結表");
                string PaoFileName = $"{MainServer.globalCommond.OperationCode_String}輕軌候車站營收班結紀錄表{item.CodeName}-{dateTimePicker2.Value.ToString("MMdd")}.xlsm";  // 原本兩班制
                //string PaoFileName_2 = $"{MainServer.globalCommond.OperationLine_String}輕軌候車站營收班結記錄表{item.CodeName}-{dateTimePicker2.Value.ToString("MMdd")}.xlsm"; // 2024-03月 改3班制
                string PaoPath = MyCommond.Path_Program + MyCommond.Path_Analyze;
                string FullPaoPath = PaoPath + PaoFileName;

                int excelCount = 0;
                ChangFileNameString:
                //FullPaoPath = PaoPath + ((excelCount == 0) ? PaoFileName : PaoFileName_2);
                MyCommond.WriteLog(ThisReceive, $"班結表路徑:{FullPaoPath}");

                if (!MyCommond.CheckFile(FullPaoPath)) { continue; }

                Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(FullPaoPath);
                NeedExport = true;
                EachSellReport PaoSale = GetPaoSale(xlWorkBook, item.CodeName);  //暫存單個Pao銷售紀錄

                List<GroupTicketList> GroupTemp = new List<GroupTicketList>();   //存放團體票部分資料

                if (!(Convert.ToInt32(PaoSale.Ticket_Group) > 0)) goto EndRead;
                GroupTemp = GetGroupTickets(xlWorkBook, item.CodeName);
                foreach (var GruopTicketTemp in GroupTemp) { GroupList.Add(GruopTicketTemp); }

            EndRead:
                xlWorkBook.Close(SaveChanges:false) ;
                PaoSaleList.Add(PaoSale);

                PaoTicketSaleList tempData = new PaoTicketSaleList()
                {
                    Operation /*           */ = PaoSale.Operation,
                    Station /*             */ = PaoSale.Station,
                    Ticket_OneDay /*       */ = PaoSale.Ticket_OneDay,
                    Ticket10 /*            */ = PaoSale.Ticket_10,
                    Ticket12 /*            */ = PaoSale.Ticket_12,
                    Ticket15 /*            */ = PaoSale.Ticket_15,
                    Ticket20 /*            */ = PaoSale.Ticket_20,
                    Ticket25 /*            */ = PaoSale.Ticket_25,
                    Ticket30 /*            */ = PaoSale.Ticket_30,
                    TicketBike /*          */ = PaoSale.Ticket_bike,
                    OpenTrafficSetTicket /**/ = PaoSale.OpenTrafficSetTicket,
                    TPASS /*               */ = PaoSale.TPass,
                    GroupCount /*          */ = ((GroupTemp == null) ? "0" : $"{GroupTemp.Count}"),
                    GroupPeople /*         */ = ((GroupTemp == null) ? "0" : $"{GroupTemp.Sum(x => Convert.ToInt16(x.People))}"),
                    GroupAmt /*            */ = ((GroupTemp == null) ? "0" : $"{GroupTemp.Sum(x => Convert.ToInt16(x.TotleAmount))}"),
                    Lanwair1 /*            */ = "0",
                    Lanwair2 /*      */ = "0",
                    AfcTakeOut /*          */ = PaoSale.AfcTackOut,
                    RepairTicket /*        */ = PaoSale.Repair,
                    ExchangeTicket20 /*    */ = "0",
                    ExchangeTicket25 /*    */ = "0",
                    ExchangeTicket30 /*    */ = "0",
                    KKDay /*               */ = "0",
                    Ntpc /*                */ = "0",
                    SplitTicker /*         */ = "0",
                    ExchangeOneDay /*      */ = "0",
                    ConcessionRefund /*    */ = PaoSale.ConcessionRefund,
                    RefundNotice /*        */ = PaoSale.RefundNotice,
                    VavmRefund /*          */ = PaoSale.VavmRefund,
                    PvRefund /*            */ = PaoSale.PvRefund,
                    CreditCardRefund /*    */ = PaoSale.CreditCardRefund,
                    QrCodeRefund /*        */ = PaoSale.QrCodeRefund,
                    SaleCardRefund /*      */ = PaoSale.SaleCardRefund,
                    GroupRefund /*         */ = PaoSale.GroupRefund
                };
                Form_PaoText_Data.Add(tempData);
            }

            xlApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

            //輸出檔案 (PaoData.csv、GroupTicket.csv)
            List<string> tempDate = (from p in GroupList group p by new { p.UseDate } into g select g.Key.UseDate).ToList();
            try
            {
                foreach (var item in tempDate)
                {
                    var tempData = (from p in GroupList where p.UseDate == item select p).ToList();
                    var date1 = Convert.ToInt32(item.Substring(0, 3)) + 1911;
                    int aa = item.Length - 3;
                    var date2 = item.Substring(3, aa);

                    MyCommond.ExportStuctData(ThisReceive, new ExportConfig()
                    {
                        Path1 = MainServer.Path_PaoData,
                        DTime = new DateTime(),
                        FirstName = $"{date1}{date2}",
                        FileName = MainServer.FileName_GroupTicket,
                        SubName = "",
                        OpLine = MainServer.globalCommond.OperationCode_String,
                        LastName = "csv",
                    }, new GroupTicketList().ExportTitle_zhTW(), tempData, Encoding.UTF8);
                }
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, MyCommond.ExportErrorMessageToLog(ex, $@"團體票異常").Message);
            }

            
        EndSub:
            MyCommond.WriteLog(ThisReceive, $"自動載入班結表結束");
        }

        private EachSellReport GetPaoSale(Excel.Workbook ExcelBook, string StationCode)
        {
            MyCommond.WriteLog(ThisReceive, $"讀取PAO總計資料");
            //var paotitle = (new EachSellReport()).ExportTitle_zhTW().Split(',');
            var List_paotitle = MainServer.globalCommond.Stock_PaoCompareTitle;
            string[] paotitle = new string[List_paotitle.Count];
            for (int i = 0; i < List_paotitle.Count; i++)
            {
                paotitle[i] = List_paotitle[i].S2.ToString();
            }
            

            Excel.Worksheet xlWorkSheet = (Excel.Worksheet)ExcelBook.Worksheets.get_Item("總計");
            Excel.Range range = xlWorkSheet.UsedRange;
            EachSellReport eachSellReport = new EachSellReport() { Operation = dateTimePicker2.Value.ToString("yyyyMMdd"), Station = StationCode };
            Type type = typeof(EachSellReport);
            for (int i_R = 1; i_R < 50; i_R++)
            {
                int Empty_Count = 0;
                for (int i_C = 1; i_C < 15; i_C++)
                {
                    int Offect_After_Empty = 0;
                    var Column_Title = Convert.ToString((range[i_R, i_C] as Excel.Range).Value2);
                    var CountTitle = 0;

                    if (Column_Title == null) continue;
                    foreach (var item in paotitle)
                    {
                        if (Column_Title == item) goto KeepGo_1;
                        CountTitle++;
                    }
                    Empty_Count++;
                    continue;

                KeepGo_1:
                    try
                    {
                        if (CountTitle > 12) { Offect_After_Empty++; }
                        var checkValue = type.GetProperties()[CountTitle].GetValue(eachSellReport);
                        var ExcelValue = (range[i_R, i_C + 1 + Offect_After_Empty] as Excel.Range).Value2;

                        if (checkValue == null)
                        {
                            type.GetProperties()[CountTitle].SetValue(eachSellReport, Convert.ToString(ExcelValue), null);
                        }
                        else
                        {
                            var mAdd = Convert.ToInt16(checkValue) + Convert.ToInt16(ExcelValue);
                            type.GetProperties()[CountTitle].SetValue(eachSellReport, Convert.ToString(mAdd), null);
                        }

                        MyCommond.WriteLog(ThisReceive, $"ExcelValue:{ExcelValue,4}");
                        MyCommond.WriteLog(ThisReceive, $"checkValue:{checkValue,4}");
                    }
                    catch (Exception ex)
                    {
                        string msg = $"活頁簿:{ExcelBook.Name}\nRow:{i_R,4},Col:{i_C + 1 + Offect_After_Empty,4}\n錯誤訊息:{ex.Message}";
                        string title = $"錯誤";
                        MyCommond.MsgShow(msg, title, MessageBoxButtons.OK, MessageBoxIcon.Error);

                        MyCommond.WriteLog(ThisReceive, $"{title}{msg}");
                        type.GetProperties()[CountTitle].SetValue(eachSellReport, Convert.ToString(0), null);
                    }
                }
                //if (Empty_Second && Empty_Count > 7) Offect_After_Empty++;
            }

            return eachSellReport;
        }

        private List<GroupTicketList> GetGroupTickets(Excel.Workbook ExcelBook, string StationCode)
        {
            MyCommond.WriteLog(ThisReceive, $"準備取得團體票資料");

            #region 變數
            MyCommond.WriteLog(ThisReceive, $"初始化變數");
            Excel.Worksheet xlWorkGroupAm = (Excel.Worksheet)ExcelBook.Worksheets.get_Item("(早)團體票");
            Excel.Range range_am = xlWorkGroupAm.UsedRange;

            Excel.Worksheet xlWorkGroupPm = (Excel.Worksheet)ExcelBook.Worksheets.get_Item("(午)團體票");
            Excel.Range range_pm = xlWorkGroupPm.UsedRange;

            Excel.Worksheet xlWorkGroupNm = (Excel.Worksheet)ExcelBook.Worksheets.get_Item("(夜)團體票");
            Excel.Range range_nm = xlWorkGroupNm.UsedRange;

            List<GroupTicketList> TempList = new List<GroupTicketList>();
            Type type = typeof(GroupTicketList);

            string[] GroupListTitle = (new GroupTicketList().ExportTitle_zhTW()).Split(',');    //取得本List的各欄位名稱
            int[] tempDataTile = new int[GroupListTitle.Length];                                //儲存資料的欄位位置

            bool FindTitle = false;
            bool jumpEx_Am = true;
            bool jumpEx_Pm = true;
            bool jumpEx_Nm = true;
            bool Finish_Group_Am = false;
            bool Finish_Group_Pm = false;
            bool Finish_Group_Nm = false;
            #endregion

            for (int ir = 1; ir < 25; ir++)
            {
                GroupTicketList GroupData_Am = new GroupTicketList() { SaleDate = dateTimePicker2.Value.ToString("yyyyMMdd"), SaleStation = StationCode };
                GroupTicketList GroupData_Pm = new GroupTicketList() { SaleDate = dateTimePicker2.Value.ToString("yyyyMMdd"), SaleStation = StationCode };
                GroupTicketList GroupData_Nm = new GroupTicketList() { SaleDate = dateTimePicker2.Value.ToString("yyyyMMdd"), SaleStation = StationCode };

                for (int ic = 1; ic < GroupListTitle.Length; ic++)
                {
                    var Data_Am = Convert.ToString((range_am[ir, ic] as Excel.Range).Value2);
                    var Data_Pm = Convert.ToString((range_pm[ir, ic] as Excel.Range).Value2);
                    var Data_Nm = Convert.ToString((range_nm[ir, ic] as Excel.Range).Value2);

                    if (!FindTitle)
                    {
                        for (int j = 0; j < GroupListTitle.Length; j++)
                        {
                            if (Data_Am == GroupListTitle[j])
                            {
                                tempDataTile[ic] = j;
                                MyCommond.WriteLog(ThisReceive, $"{j,4}. 找到標頭寫入{ic,4}");
                                break;
                            }
                        }
                    }
                    else
                    {
                        if (ic == 1) continue;
                        int e_column = tempDataTile[ic];
                        try
                        {
                            if (!jumpEx_Am)
                            {
                                //if (ic < 8 && Data_Am == "") { Finish_Group_Am = true; }

                                if (!Finish_Group_Am)
                                {
                                    if (ic < 8 && Data_Am == "") { Finish_Group_Am = true; MyCommond.WriteLog(ThisReceive, $"早班團體票已搜索完畢"); }
                                    else
                                    {
                                        var checkValue = type.GetProperties()[e_column].GetValue(GroupData_Am);
                                        if (checkValue == null)
                                        {
                                            type.GetProperties()[e_column].SetValue(GroupData_Am, Convert.ToString(Data_Am), null);
                                        }
                                        else
                                        {
                                            var mAdd = Convert.ToInt16(checkValue) + Convert.ToInt16(Data_Am);
                                            type.GetProperties()[e_column].SetValue(GroupData_Am, Convert.ToString(mAdd), null);
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"ERROR! \n{MyCommond.ExportErrorMessageToLog(ex).Message}");
                        } //早

                        try
                        {
                            if (!jumpEx_Pm)
                            {
                                //if (ic < 8 && Data_Pm == "") { Finish_Group_Pm = true; }

                                if (!Finish_Group_Pm)
                                {
                                    if (ic < 8 && Data_Pm == "") { Finish_Group_Pm = true; MyCommond.WriteLog(ThisReceive, $"午班團體票已搜索完畢"); }
                                    else
                                    {
                                        var checkValue = type.GetProperties()[e_column].GetValue(GroupData_Pm);
                                        if (checkValue == null)
                                        {
                                            type.GetProperties()[e_column].SetValue(GroupData_Pm, Convert.ToString(Data_Pm), null);
                                        }
                                        else
                                        {
                                            var mAdd = Convert.ToInt16(checkValue) + Convert.ToInt16(Data_Pm);
                                            type.GetProperties()[e_column].SetValue(GroupData_Pm, Convert.ToString(mAdd), null);
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"ERROR! \n{MyCommond.ExportErrorMessageToLog(ex).Message}");
                        } //午

                        try
                        {
                            if (!jumpEx_Nm)
                            {
                                //if (ic < 8 && Data_Nm == "") { Finish_Group_Nm = true; }

                                if (!Finish_Group_Nm)
                                {
                                    if (ic < 8 && Data_Nm == "") { Finish_Group_Nm = true; MyCommond.WriteLog(ThisReceive, $"夜班團體票已搜索完畢"); }
                                    else
                                    {
                                        var checkValue = type.GetProperties()[e_column].GetValue(GroupData_Nm);
                                        if (checkValue == null)
                                        {
                                            type.GetProperties()[e_column].SetValue(GroupData_Nm, Convert.ToString(Data_Nm), null);
                                        }
                                        else
                                        {
                                            var mAdd = Convert.ToInt16(checkValue) + Convert.ToInt16(Data_Nm);
                                            type.GetProperties()[e_column].SetValue(GroupData_Nm, Convert.ToString(mAdd), null);
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"ERROR! \n{MyCommond.ExportErrorMessageToLog(ex).Message}");
                        } //夜

                    }
                }

                if (GroupData_Am.Entry != "" && GroupData_Am.Entry != null) { TempList.Add(GroupData_Am); }
                if (GroupData_Pm.Entry != "" && GroupData_Pm.Entry != null) { TempList.Add(GroupData_Pm); }
                if (GroupData_Nm.Entry != "" && GroupData_Nm.Entry != null) { TempList.Add(GroupData_Nm); }

                if (FindTitle) { jumpEx_Am = false; jumpEx_Pm = false; jumpEx_Nm = false; }
                if (FindTitle) { continue; }

                MyCommond.WriteLog(ThisReceive, $"檢查搜索到的標頭與設定的標頭是否一致");
                int temp_i = 0;
                foreach (var item in tempDataTile) { if (item != 0) { temp_i++; } }
                FindTitle = temp_i == tempDataTile.Length - 2;
                MyCommond.WriteLog(ThisReceive, $"FindTitle:{((FindTitle) ? "相同" : "不同")}");
            }

            foreach (var item in TempList)
            {
                int y = Convert.ToInt16(item.UseDate.Substring(0, 3));
                int m = Convert.ToInt16(item.UseDate.Substring(3, 2));
                int d = Convert.ToInt16(item.UseDate.Substring(5, 2));
                int h = Convert.ToInt16(item.SaleTime.Substring(0, 2));
                int m2 = Convert.ToInt16(item.SaleTime.Substring(3, 2));
                var date1 = new DateTime(((y < 1911)? (y + 1911) : y), m, d, h, m2, 0);
                if (date1 < dateTimePicker1.Value)
                {
                    string[] arrayStr1 = new GroupTicketList().ExportTitle_zhTW().Split(',');
                    string[] arrayStr2 = item.ToString().Split(',');
                    string errorStr = "";
                    for (int i = 0; i < arrayStr1.GetLength(0); i++)
                    {
                        errorStr += ((errorStr != "") ? $"\n" : "") + $@"{arrayStr1[i]}:{arrayStr2[i]}";
                    }
                    MyCommond.WriteLog(ThisReceive,
                        MyCommond.ExportErrorMessageToLog(new Exception(errorStr), "使用日期可能有誤").Message
                    );
                }
            } // 檢測使用日期有無異常

            MyCommond.WriteLog(ThisReceive, $"從班結表取出{TempList.Count}筆團體票資料");

            return TempList;
        }

        private bool CheckAndExportData<TResult>(ExportConfig ec, string DataTitle, List<TResult> Data, Encoding Ecode) where TResult : class, new()
        {
            string[] TextWay = new string[] { "csv", "txt", "config" };
            string[] ExcelWay = new string[] { "xls", "xlsx" };
            string FullPath = $"{ec.Path1}/{ec.FirstName}_{ec.FileName}_{ec.OpLine}.{ec.LastName}";

            bool HaveFile = MyCommond.CheckFile(FullPath);
            List<TResult> TempList = new List<TResult>();

            if (HaveFile)
            {
                // 要修改檢查內容有沒有相同的
                var sss = MyCommond.ToViewModel<TResult>(ThisReceive, FullPath, DataTitle, 1);

                foreach (var item in Data)
                {
                    bool HaveData_b = false;
                    foreach (var jtem in sss)
                    {
                        if (item.ToString() == jtem.ToString())
                        {
                            HaveData_b = true;
                            break;
                        }
                    }
                    if (!HaveData_b)
                    {
                        TempList.Add(item);
                    }
                }
            }
            else
            {
                TempList = Data.ToList();
            }

            if (TempList.Count == 0)
            {
                MyCommond.WriteLog(ThisReceive, $"無可輸出資料");
                return true;
            }

            if (TextWay.Contains(ec.LastName))
            {
                MyCommond.WriteLog(ThisReceive, $"開始輸出文字檔");
                if (!HaveFile && DataTitle != "")
                {
                    if (!File.Exists(FullPath)) File.WriteAllText(FullPath, DataTitle + Environment.NewLine, Ecode);
                    else File.AppendAllText(FullPath, DataTitle + Environment.NewLine, Ecode);
                }

                foreach (var item in TempList)
                {
                    if (!File.Exists(FullPath)) File.WriteAllText(FullPath, item.ToString() + Environment.NewLine, Ecode);
                    else File.AppendAllText(FullPath, item.ToString() + Environment.NewLine, Ecode);
                }
            }
            else if (ExcelWay.Contains(ec.LastName))
            {
                MyCommond.WriteLog(ThisReceive, $"尚未提供試算表輸出");
                //WriteLog(ThisReceive, $"開始輸出試算表");


            }
            else
            {
                MyCommond.WriteLog(ThisReceive, $"未發現正確附檔名 -> {ec.LastName}");
            }

            return false;
        }

        #endregion

        #endregion

        private void EmptyMonthen(object sender, EventArgs e)
        {

        }

        #region Timer

        private void Timer_Start()
        {
            bool checkStart = false;
            while (!checkStart)
            {
                checkStart = MainServer.IsRunning;
                System.Threading.Thread.Sleep(1000);
            }

            if (!true)
            {
                //Text_timer.Start();
                return;
            }

            List<CheckBoxs> _CB = MyCheckBoxs.FindAll(x => x.Status == true).ToList();
            foreach (var item in _CB)
            {
                MyCommond.WriteLog(ThisReceive, $"Timer - {item.Methode.ToDescription()} 處理");
                int Timer_Search_i = MyTimers.FindIndex(x => x.Methode == item.Methode);
                if (Timer_Search_i == -1)
                {
                    MyCommond.WriteLog(ThisReceive, $"沒找到相符的線程");
                    continue;
                }

                MyCommond.WriteLog(ThisReceive, $"找到Timer");
                Timers timers = MyTimers.Find(x => x.Methode == item.Methode);
                if (timers.Status == StatusEnum.Start)
                {
                    MyCommond.WriteLog(ThisReceive, $"已經是啟動狀態");
                    continue;
                }
                timers.Status = StatusEnum.Start;
                //timers.Item.Start();
                MyCommond.WriteLog(ThisReceive, $"線程啟動");
            }
        }
        string timerLock = "";
        private void TimerMethode(LrtTxnClientSwitch Methode)
        {
            lock (timerLock)
            {
                Timers _Timer = MyTimers.Find(x => x.Methode == Methode);
                if (_Timer.Status == StatusEnum.Stop) return;
                int Max = MainServer.LrtClients.Find(x => x.MethodeName == Methode).MaxCount;
                int Min = MainServer.LrtClients.Find(x => x.MethodeName == Methode).MinCount;
                int Value = MainServer.LrtClients.Find(x => x.MethodeName == Methode).ValueCount;
                MyCommond.WriteLog(ThisReceive, $"更新 {Methode.ToDescription(),10} 狀態 Max:{Max,4}, Min:{Min,4}, Value:{Value,4}");
                this.MyCommond.InvokeIfRequired(this, () =>
                {
                    MyProgressBars.Find(x => x.Methode == Methode).Item.Maximum = Max;
                    MyProgressBars.Find(x => x.Methode == Methode).Item.Minimum = Min;
                    MyProgressBars.Find(x => x.Methode == Methode).Item.Value = Value;
                });
                if (Max == Value)
                {
                    _Timer.Status = StatusEnum.Stop;
                    //_Timer.Item.Stop();
                }
                bool Button_Eabled = MyTimers.FindAll(x => x.Status == StatusEnum.Start).Count() > 0;
                bool Client_Status = MainServer.LrtClients.FindAll(x => x.ThisMethode == StatusEnum.Finish).Count == MyCheckBoxs.Count();
                if (Button_Eabled || Client_Status) return;
                MainServer.LrtClients.Remove(MainServer.LrtClients.Find(x => x.MethodeName == Methode));
                All_Item_Status(true);
            }
        }

        #region Q

        private void TextTime(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"???");
            List<CheckBoxs> _CB = MyCheckBoxs.FindAll(x => x.Status == true).ToList();
            foreach (var item in _CB)
            {
                LrtTxnClientSwitch methode = item.Methode;
                //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
                TimerMethode(methode);
            }
        }

        private void Timer1(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)1;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer2(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)2;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer3(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)3;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer4(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)4;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer5(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)5;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer6(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)6;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer7(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)7;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer8(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)8;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer9(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)9;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer10(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)10;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer11(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)11;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer12(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)12;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer13(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)13;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer14(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)14;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer15(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)15;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer16(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)16;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer17(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)17;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer18(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)18;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer19(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)19;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        private void Timer20(object sender, EventArgs e)
        {
            LrtTxnClientSwitch methode = (LrtTxnClientSwitch)20;
            //MyCommond.WriteLog(ThisReceive, $"{methode.ToString()} - 啟動");
            TimerMethode(methode);
        }

        #endregion

        #endregion

        #region Form - UI

        #region Control

        #region Text

        public void LoadPaoLable(string str) { MyCommond.InvokeIfRequired(this, () => { label_Loading_Pao.Text = str; }); }

        public void txt_CreditCardPath(string path) { MyCommond.InvokeIfRequired(this, () => { textBox_CreditCard.Text = path; }); }

        public void txt_QRCodePath(string path) { MyCommond.InvokeIfRequired(this, () => { textBox_QRCode.Text = path; }); }

        public void Data2Form(List<PaoTicketSaleList> PaoData)
        {
            MyCommond.WriteLog(ThisReceive, $"資料轉換成畫面");
            this.MyCommond.InvokeIfRequired(this, () =>
            {

                if (PaoData.Count() > 0)
                {

                    #region Group Pao+ 1

                    textBox_OneDay_1.Text = PaoData[0].Ticket_OneDay;
                    textBox_OneDay_E_1.Text = "0";
                    textBox_Ticket_10_1.Text = PaoData[0].Ticket10;
                    textBox_Ticket_12_1.Text = PaoData[0].Ticket12;
                    textBox_Ticket_15_1.Text = PaoData[0].Ticket15;
                    textBox_Ticket_20_1.Text = PaoData[0].Ticket20;
                    textBox_Ticket_25_1.Text = PaoData[0].Ticket25;
                    textBox_Ticket_30_1.Text = PaoData[0].Ticket30;
                    textBox_Ticket_50_1.Text = PaoData[0].TicketBike;
                    textBox_OpenTransaction_1.Text = PaoData[0].OpenTrafficSetTicket;
                    textBox_Tpass_1.Text = PaoData[0].TPASS;
                    textBox_Group_Count_1.Text = PaoData[0].GroupCount;
                    textBox_Group_People_1.Text = PaoData[0].GroupPeople;
                    textBox_Group_Amount_1.Text = PaoData[0].GroupAmt;
                    textBox_Lanwair1_1.Text = PaoData[0].Lanwair1;
                    textBox_Lanwair2_1.Text = PaoData[0].Lanwair2;
                    textBox_VavmTake_1.Text = PaoData[0].AfcTakeOut;
                    textBox_Repair_1.Text = PaoData[0].RepairTicket;

                    #endregion

                    #region Group Pao- 1

                    textBox_ConcessionRefund_1.Text = PaoData[0].ConcessionRefund;
                    textBox_RefundNotice_1.Text = PaoData[0].RefundNotice;
                    textBox_VavmRefund_1.Text = PaoData[0].VavmRefund;
                    textBox_PvRefund_1.Text = PaoData[0].PvRefund;
                    textBox_CreditCardRefund_1.Text = PaoData[0].CreditCardRefund;
                    textBox_QrCodeRefund_1.Text = PaoData[0].QrCodeRefund;
                    textBox_SaleCardRefund_1.Text = PaoData[0].SaleCardRefund;
                    textBox_GroupRefund_1.Text = PaoData[0].GroupRefund;

                    #endregion

                    #region Group Exchange Pao 1

                    textBox_ExchangeTicket_20_1.Text = PaoData[0].ExchangeTicket20;
                    textBox_ExchangeTicket_25_1.Text = PaoData[0].ExchangeTicket25;
                    textBox_ExchangeTicket_30_1.Text = PaoData[0].ExchangeTicket30;

                    #endregion

                    #region Group OneDay

                    textBox_ExchangeOneDay_1.Text = PaoData[0].ExchangeOneDay;
                    textBox_KKDay.Text = PaoData[0].KKDay;
                    textBox_Ntpc.Text = PaoData[0].Ntpc;
                    textBox_SplitTicker.Text = PaoData[0].SplitTicker;

                    #endregion

                }

                if (PaoData.Count() > 1)
                {

                    #region Group Pao+ 2

                    textBox_OneDay_2.Text = PaoData[1].Ticket_OneDay;
                    textBox_OneDay_E_2.Text = "0";
                    textBox_Ticket_10_2.Text = PaoData[1].Ticket10;
                    textBox_Ticket_12_2.Text = PaoData[1].Ticket12;
                    textBox_Ticket_15_2.Text = PaoData[1].Ticket15;
                    textBox_Ticket_20_2.Text = PaoData[1].Ticket20;
                    textBox_Ticket_25_2.Text = PaoData[1].Ticket25;
                    textBox_Ticket_30_2.Text = PaoData[1].Ticket30;
                    textBox_Ticket_50_2.Text = PaoData[1].TicketBike;
                    textBox_OpenTransaction_2.Text = PaoData[1].OpenTrafficSetTicket;
                    textBox_Tpass_2.Text = PaoData[1].TPASS;
                    textBox_Group_Count_2.Text = PaoData[1].GroupCount;
                    textBox_Group_People_2.Text = PaoData[1].GroupPeople;
                    textBox_Group_Amount_2.Text = PaoData[1].GroupAmt;
                    textBox_Lanwair1_2.Text = PaoData[1].Lanwair1;
                    textBox_Lanwair2_2.Text = PaoData[1].Lanwair2;
                    textBox_VavmTake_2.Text = PaoData[1].AfcTakeOut;
                    textBox_Repair_2.Text = PaoData[1].RepairTicket;

                    #endregion

                    #region Group Pao- 2

                    textBox_ConcessionRefund_2.Text = PaoData[1].ConcessionRefund;
                    textBox_RefundNotice_2.Text = PaoData[1].RefundNotice;
                    textBox_VavmRefund_2.Text = PaoData[1].VavmRefund;
                    textBox_PvRefund_2.Text = PaoData[1].PvRefund;
                    textBox_CreditCardRefund_2.Text = PaoData[1].CreditCardRefund;
                    textBox_QrCodeRefund_2.Text = PaoData[1].QrCodeRefund;
                    textBox_SaleCardRefund_2.Text = PaoData[1].SaleCardRefund;
                    textBox_GroupRefund_2.Text = PaoData[1].GroupRefund;

                    #endregion

                    #region Group Exchange Pao 2

                    textBox_ExchangeTicket_20_2.Text = PaoData[1].ExchangeTicket20;
                    textBox_ExchangeTicket_25_2.Text = PaoData[1].ExchangeTicket25;
                    textBox_ExchangeTicket_30_2.Text = PaoData[1].ExchangeTicket30;

                    #endregion

                    #region Group OneDay

                    textBox_ExchangeOneDay_2.Text = PaoData[1].ExchangeOneDay;

                    #endregion

                }

            });

        }

        public List<PaoTicketSaleList> Form2Data()
        {
            MyCommond.WriteLog(ThisReceive, $"畫面轉換成資料");
            List<PaoTicketSaleList> PaoData = new List<PaoTicketSaleList>();

            var ec = new ExportConfig()
            {
                Path1 = MainServer.Path_PaoData,
                DTime = new DateTime(),
                FirstName = dateTimePicker2.Value.ToString("yyyyMM"),
                FileName = MainServer.FileName_PaoCsv,
                SubName = "",
                OpLine = MainServer.globalCommond.OperationCode_String,
                LastName = "csv",
            };
            string PaoDataPath = ec.ToString() + ec.LastName;
            var PaoDataList = MyCommond.ToViewModel<PaoTicketSaleList>(ThisReceive, PaoDataPath, (new PaoTicketSaleList()).ExportTitle_zhTW(), 1);

            var searchPaoData = PaoDataList.FindAll(x => x.Operation == dateTimePicker2.Value.ToString("yyyyMMdd")).ToList();

            if (searchPaoData.Count > 0) return null;

            this.MyCommond.InvokeIfRequired(this, () =>
            {
                var allStation = MainServer.globalCommond.Stock_Lrt_Station.FindAll(x => x.StationPao == "1");
                int PaoCount = allStation.Count();

                if (PaoCount > 0)
                {
                    PaoTicketSaleList paoticket = new PaoTicketSaleList();
                    paoticket.Operation = dateTimePicker2.Value.ToString("yyyyMMdd");
                    paoticket.Station = allStation[0].CodeName;
                    #region Group Pao+ 1

                    paoticket.Ticket_OneDay = textBox_OneDay_1.Text;
                    //textBox_OneDay_E_1.Text = "0";
                    paoticket.Ticket10 = textBox_Ticket_10_1.Text;
                    paoticket.Ticket12 = textBox_Ticket_12_1.Text;
                    paoticket.Ticket15 = textBox_Ticket_15_1.Text;
                    paoticket.Ticket20 = textBox_Ticket_20_1.Text;
                    paoticket.Ticket25 = textBox_Ticket_25_1.Text;
                    paoticket.Ticket30 = textBox_Ticket_30_1.Text;
                    paoticket.TicketBike = textBox_Ticket_50_1.Text;
                    paoticket.OpenTrafficSetTicket = textBox_OpenTransaction_1.Text;
                    paoticket.TPASS = textBox_Tpass_1.Text;
                    paoticket.GroupCount = textBox_Group_Count_1.Text;
                    paoticket.GroupPeople = textBox_Group_People_1.Text;
                    paoticket.GroupAmt = textBox_Group_Amount_1.Text;
                    paoticket.Lanwair1 = textBox_Lanwair1_1.Text;
                    paoticket.Lanwair2 = textBox_Lanwair2_1.Text;
                    paoticket.AfcTakeOut = textBox_VavmTake_1.Text;
                    paoticket.RepairTicket = textBox_Repair_1.Text;

                    #endregion

                    #region Group Pao- 1

                    paoticket.ConcessionRefund = textBox_ConcessionRefund_1.Text;
                    paoticket.RefundNotice = textBox_RefundNotice_1.Text;
                    paoticket.VavmRefund = textBox_VavmRefund_1.Text;
                    paoticket.PvRefund = textBox_PvRefund_1.Text;
                    paoticket.CreditCardRefund = textBox_CreditCardRefund_1.Text;
                    paoticket.QrCodeRefund = textBox_QrCodeRefund_1.Text;
                    paoticket.SaleCardRefund = textBox_SaleCardRefund_1.Text;
                    paoticket.GroupRefund = textBox_GroupRefund_1.Text;

                    #endregion

                    #region Group Exchange Pao 1

                    paoticket.ExchangeTicket20 = textBox_ExchangeTicket_20_1.Text;
                    paoticket.ExchangeTicket25 = textBox_ExchangeTicket_25_1.Text;
                    paoticket.ExchangeTicket30 = textBox_ExchangeTicket_30_1.Text;

                    #endregion

                    #region Group OneDay
                    paoticket.ExchangeOneDay = textBox_ExchangeOneDay_1.Text;
                    paoticket.KKDay = textBox_KKDay.Text;
                    paoticket.Ntpc = textBox_Ntpc.Text;
                    paoticket.SplitTicker = textBox_SplitTicker.Text;

                    #endregion

                    PaoData.Add(paoticket);
                }

                if (PaoCount > 1)
                {
                    PaoTicketSaleList paoticket = new PaoTicketSaleList();
                    paoticket.Operation = dateTimePicker2.Value.ToString("yyyyMMdd");
                    paoticket.Station = allStation[1].CodeName;
                    #region Group Pao+ 2

                    paoticket.Ticket_OneDay = textBox_OneDay_2.Text;
                    //textBox_OneDay_E_2.Text = "0";
                    paoticket.Ticket10 = textBox_Ticket_10_2.Text;
                    paoticket.Ticket12 = textBox_Ticket_12_2.Text;
                    paoticket.Ticket15 = textBox_Ticket_15_2.Text;
                    paoticket.Ticket20 = textBox_Ticket_20_2.Text;
                    paoticket.Ticket25 = textBox_Ticket_25_2.Text;
                    paoticket.Ticket30 = textBox_Ticket_30_2.Text;
                    paoticket.TicketBike = textBox_Ticket_50_2.Text;
                    paoticket.OpenTrafficSetTicket = textBox_OpenTransaction_2.Text;
                    paoticket.TPASS = textBox_Tpass_2.Text;
                    paoticket.GroupCount = textBox_Group_Count_2.Text;
                    paoticket.GroupPeople = textBox_Group_People_2.Text;
                    paoticket.GroupAmt = textBox_Group_Amount_2.Text;
                    paoticket.Lanwair1 = textBox_Lanwair1_2.Text;
                    paoticket.Lanwair2 = textBox_Lanwair2_2.Text;
                    paoticket.AfcTakeOut = textBox_VavmTake_2.Text;
                    paoticket.RepairTicket = textBox_Repair_2.Text;

                    #endregion

                    #region Group Pao- 2

                    paoticket.ConcessionRefund = textBox_ConcessionRefund_2.Text;
                    paoticket.RefundNotice = textBox_RefundNotice_2.Text;
                    paoticket.VavmRefund = textBox_VavmRefund_2.Text;
                    paoticket.PvRefund = textBox_PvRefund_2.Text;
                    paoticket.CreditCardRefund = textBox_CreditCardRefund_2.Text;
                    paoticket.QrCodeRefund = textBox_QrCodeRefund_2.Text;
                    paoticket.SaleCardRefund = textBox_SaleCardRefund_2.Text;
                    paoticket.GroupRefund = textBox_GroupRefund_2.Text;

                    #endregion

                    #region Group Exchange Pao 2

                    paoticket.ExchangeTicket20 = textBox_ExchangeTicket_20_2.Text;
                    paoticket.ExchangeTicket25 = textBox_ExchangeTicket_25_2.Text;
                    paoticket.ExchangeTicket30 = textBox_ExchangeTicket_30_2.Text;

                    #endregion

                    #region Group OneDay

                    paoticket.ExchangeOneDay = textBox_ExchangeOneDay_2.Text;

                    #endregion

                    PaoData.Add(paoticket);
                }
            });

            return PaoData;
        }

        #endregion

        #region CheckBox

        bool checkBatch = false;
        private CheckBox Check_checkBox(LrtTxnClientSwitch Methode)
        {
            MyCommond.WriteLog(ThisReceive, $"線程對應核取方塊");
            switch ((int)Methode)
            {
                case 1: return checkBox_G_1_1;
                case 2: return checkBox_G_1_2;
                case 3: return checkBox_G_1_3;
                case 4: return checkBox_G_1_4;
                //case 5: return checkBox_G_1_5; 
                case 6: return checkBox_G_1_6;
                case 7: return checkBox_G_1_7;
                case 8: return checkBox_G_1_8;
                //case 9 : return checkBox_G_2_1; 
                case 10: return checkBox_G_2_2;
                case 11: return checkBox_G_2_3;
                case 12: return checkBox_G_2_4;
                case 13: return checkBox_G_2_5;
                case 14: return checkBox_G_2_6;
                case 15: return checkBox_G_2_7;
                case 16: return checkBox_G_2_8;
                case 17: return checkBox_G_2_9;
                case 18: return checkBox_G_2_10;
                case 19: return checkBox_G_2_11;
                case 20: return checkBox_G_2_12;
                default: return null;
            }
        }

        private LrtTxnClientSwitch Check_ClientSwitch(CheckBox checkBox)
        {
            MyCommond.WriteLog(ThisReceive, $"核取方塊對應線程");
            return MyCheckBoxs.Find(x => x.Item.Text == checkBox.Text).Methode;
        }

        // 單選用
        private void checkIO(CheckBox checkBox)
        {
            checkBatch = true;
            LrtTxnClientSwitch Methode = Check_ClientSwitch(checkBox);
            MyCommond.WriteLog(ThisReceive, $"檢查線程 {Methode.ToString()} - {((checkBox.Checked) ? "啟動" : "關閉")}");
            int Modthen_i = (int)Methode;
            bool checkIsDay = Modthen_i < 9;    // 線程數值小於9為日報表，大於等於9的是月報表
            foreach (var item in MyCheckBoxs)
            {
                bool itemIsDay = (int)item.Methode < 9;            // 判斷這個item是不是日報表
                bool itemStatus = item.Status;                      // 取得這個item的狀態是否啟動
                bool Reverse = false;                            // 是否需要反轉
                bool IsCheck = item.Item.Text == checkBox.Text;  // checkbox與item的名稱是否相同

                /**/
                if (checkIsDay && itemIsDay) { }                     // 此核取方塊是 每日 報表 且 物件也是每日 不需要轉
                else if (checkIsDay && !itemIsDay) { Reverse = true; }     // 此核取方塊是 每日 報表 且 物件  是每月   需要轉
                else if (!checkIsDay && itemIsDay) { Reverse = true; }     // 此核取方塊是 每月 報表 且 物件  是每日   需要轉
                else if (!checkIsDay && !itemIsDay) { }                     // 此核取方塊是 每月 報表 且 物件也是每月 不需要轉

                //MyCommond.WriteLog(ThisReceive, $"" +
                //    $"線程是否為每日:{((checkIsDay) ? "是" : "否")}, " +
                //    $"物件是否為每日:{((itemIsDay) ? "是" : "否")}, " +
                //    $"是否需要反轉  :{((Reverse) ? "是" : "否")}, " +
                //    $"線程與物件相同:{((IsCheck) ? "是" : "否")}");

                if (IsCheck && Reverse)
                {
                    //MyCommond.WriteLog(ThisReceive, $"00");
                    // null
                }
                else if (IsCheck && !Reverse)
                {
                    //MyCommond.WriteLog(ThisReceive, $"01");
                    CheckBoxs checkBoxs = MyCheckBoxs.Find(x => x.Item.Text == checkBox.Text);
                    checkBoxs.Status = checkBox.Checked;
                    checkBoxs.Item = checkBox;
                }
                else if (!IsCheck && Reverse)
                {
                    // 不是核取方塊選取的且需要被取消選取
                    //MyCommond.WriteLog(ThisReceive, $"10");
                    if (!itemStatus) continue; // 這物件如果是未被選取的就跳過
                    CheckBox jtem = Check_checkBox(item.Methode);
                    jtem.Checked = !jtem.Checked;
                    CheckBoxs checkBoxs = MyCheckBoxs.Find(x => x.Item.Text == jtem.Text);
                    checkBoxs.Status = jtem.Checked;
                    checkBoxs.Item = jtem;
                }
                else if (!IsCheck && !Reverse)
                {
                    //MyCommond.WriteLog(ThisReceive, $"11");
                    // Null
                }
            }
            checkBatch = false;
        }

        // 多選用
        private void checkIO(DatePark DM, bool TF)
        {
            checkBatch = true;
            MyCommond.WriteLog(ThisReceive, $"批次選取翻轉");
            foreach (var item in MyCheckBoxs)
            {
                bool Reverse = false;
                int i = (int)item.Methode;
                if (DM == DatePark.Day)
                {
                    if (i > 8) Reverse = true;
                }
                else if (DM == DatePark.Month)
                {
                    if (i < 9) Reverse = true;
                }
                else return;

                CheckBox checkBox_1 = Check_checkBox(item.Methode);
                if (checkBox_1 == null) { goto NextOne; }
                string Nothing = "";
                if (Reverse && TF) { checkBox_1.Checked = !TF; }
                else if (Reverse && !TF) { Nothing = ""; }
                else { checkBox_1.Checked = TF; }

                LrtTxnClientSwitch Methode = Check_ClientSwitch(checkBox_1);
                MyCommond.WriteLog(ThisReceive, $"{Methode.ToString()} - {((checkBox_1.Checked) ? "啟動" : "關閉")}");
                CheckBoxs checkBoxs_2 = MyCheckBoxs.Find(x => x.Methode == item.Methode);
                checkBoxs_2.Status = checkBox_1.Checked;
                checkBoxs_2.Item = checkBox_1;
            NextOne:
                Nothing = "";
            }
            checkBatch = false;
        }

        // 按鈕用
        private void checkIO(LrtTxnClientSwitch Methode, bool TF)
        {
            checkBatch = true;
            MyCommond.WriteLog(ThisReceive, $"Button - {Methode.ToString()} 執行");
            for (int i = 1; i <= 18; i++)
            {
                CheckBoxs checkBoxs_0 = MyCheckBoxs.Find(x => x.Methode == (LrtTxnClientSwitch)i);
                if (checkBoxs_0.Methode == Methode) { checkBoxs_0.Status = TF; }
                else { checkBoxs_0.Status = !TF; }
            }
            checkBatch = false;
        }

        #endregion

        #region Status

        private void Group1_Status(string item, bool TF)
        {
            MyCommond.WriteLog(ThisReceive, $"{item} - {((TF) ? "開啟" : "關閉")}");
            MyCommond.InvokeIfRequired(this, () =>
            {
                if (item == "Thread_Use") checkBox1.Checked = TF;
                else if (item == "Excel_Applaction_Show") checkBox2.Checked = TF;
                else if (item == "Auto_Pao_Excel") checkBox3.Checked = TF;
                else if (item == "Install_MobilePay") checkBox4.Checked = TF;
            });

        }

        private void All_Item_Status(bool TF)
        {
            Pao_Status(1, TF);
            if (MainServer.globalCommond.Stock_Lrt_Station.FindAll(x => x.StationPao == "1").Count() > 1) Pao_Status(2, TF);
            //Pao_Status(2, TF); /* 測試 */

            Group1_Status(TF);
            Group2_Status(TF);
            Group3_Status(TF);
            Button_Status(TF);
            Other_Status(TF);
        }

        private void Group1_Status(bool TF)
        {
            MyCommond.WriteLog(ThisReceive, $"程式核取區塊 - {((TF) ? "開啟" : "關閉")}");
            MyCommond.InvokeIfRequired(this, () =>
            {
                //checkBox1.Enabled = TF;
                checkBox2.Enabled = TF;
                checkBox3.Enabled = TF;
                checkBox4.Enabled = TF;
            });
        }

        private void Group2_Status(bool TF)
        {
            MyCommond.WriteLog(ThisReceive, $"日運量批次區塊 - {((TF) ? "開啟" : "關閉")}");
            MyCommond.InvokeIfRequired(this, () =>
            {
                if (MainServer.globalCommond.Finish_Button >= 1) dateTimePicker3.Enabled = TF;     //開始日期
                if (MainServer.globalCommond.Finish_Button >= 1) dateTimePicker4.Enabled = TF;     //結束日期
                if (MainServer.globalCommond.Finish_Button >= 1) button_G_1_Start.Enabled = TF;    //日執行(限制運量營收一起跑)
                if (MainServer.globalCommond.Finish_Button >= 1) button_G_1_All.Enabled = TF;    //日全選
                if (MainServer.globalCommond.Finish_Button >= 1) button_G_1_Cancel.Enabled = TF;    //日重製
                if (MainServer.globalCommond.Finish_Button >= 1) checkBox_G_1_1.Enabled = TF;
                if (MainServer.globalCommond.Finish_Button >= 2) checkBox_G_1_2.Enabled = TF;
                if (MainServer.globalCommond.Finish_Button >= 3) checkBox_G_1_3.Enabled = TF;
                if (MainServer.globalCommond.Finish_Button >= 4) checkBox_G_1_4.Enabled = TF;
                //if (MainServer.globalCommonds.Finish_Button >= 5) checkBox_G_1_5.Enabled = TF;   /* 永遠上鎖 */
                if (MainServer.globalCommond.Finish_Button >= 6) checkBox_G_1_6.Enabled = TF;
                if (MainServer.globalCommond.Finish_Button >= 7) checkBox_G_1_7.Enabled = TF;
                if (MainServer.globalCommond.Finish_Button >= 8) checkBox_G_1_8.Enabled = TF;
            });
        }

        private void Group3_Status(bool TF)
        {
            MyCommond.WriteLog(ThisReceive, $"月運量批次區塊 - {((TF) ? "開啟" : "關閉")}");
            MyCommond.InvokeIfRequired(this, () =>
            {
                if (MainServer.globalCommond.Finish_Button >= 9) dateTimePicker5.Enabled = TF;     //開始日期
                if (MainServer.globalCommond.Finish_Button >= 9) dateTimePicker6.Enabled = TF;     //結束日期
                if (MainServer.globalCommond.Finish_Button >= 9) button_G_2_Start.Enabled = TF;    //月執行
                if (MainServer.globalCommond.Finish_Button >= 9) button_G_2_All.Enabled = TF;    //月全選
                if (MainServer.globalCommond.Finish_Button >= 9) button_G_2_Cancel.Enabled = TF;    //月重製
                //if (MainServer.globalCommonds.Finish_Button >= 9) checkBox_G_2_1.Enabled = TF;    /* 永遠上鎖 */
                if (MainServer.globalCommond.Finish_Button >= 10) checkBox_G_2_2.Enabled = TF;
                if (MainServer.globalCommond.Finish_Button >= 11) checkBox_G_2_3.Enabled = TF;
                if (MainServer.globalCommond.Finish_Button >= 12) checkBox_G_2_4.Enabled = TF;
                if (MainServer.globalCommond.Finish_Button >= 13) checkBox_G_2_5.Enabled = TF;
                if (MainServer.globalCommond.Finish_Button >= 14) checkBox_G_2_6.Enabled = TF;
                if (MainServer.globalCommond.Finish_Button >= 15) checkBox_G_2_7.Enabled = TF;
                if (MainServer.globalCommond.Finish_Button >= 16) checkBox_G_2_8.Enabled = TF;
                if (MainServer.globalCommond.Finish_Button >= 17) checkBox_G_2_9.Enabled = TF;
                if (MainServer.globalCommond.Finish_Button >= 18) checkBox_G_2_10.Enabled = TF;
                if (MainServer.globalCommond.Finish_Button >= 19) checkBox_G_2_11.Enabled = TF;
                if (MainServer.globalCommond.Finish_Button >= 20) checkBox_G_2_12.Enabled = TF;
            });
        }

        private void Pao_Status(int item, bool TF)
        {
            MyCommond.WriteLog(ThisReceive, $"PAO Text({((item == 1) ? "左側" : "右側")}) - {((TF) ? "開啟" : "關閉")}");
            this.MyCommond.InvokeIfRequired(this, () =>
            {
                if (item == 1)
                {

                    #region Label

                    label_pao_1_1.Enabled = TF;
                    label_pao_1_2.Enabled = TF;
                    label_pao_1_3.Enabled = TF;
                    label_pao_1_4.Enabled = TF;

                    #endregion

                    #region Group Pao+ 1

                    textBox_OneDay_1.Enabled = TF;
                    textBox_OneDay_E_1.Enabled = TF;
                    textBox_Ticket_10_1.Enabled = TF;
                    textBox_Ticket_12_1.Enabled = TF;
                    textBox_Ticket_15_1.Enabled = TF;
                    textBox_Ticket_20_1.Enabled = TF;
                    textBox_Ticket_25_1.Enabled = TF;
                    textBox_Ticket_30_1.Enabled = TF;
                    textBox_Ticket_50_1.Enabled = TF;
                    textBox_OpenTransaction_1.Enabled = TF;
                    textBox_Tpass_1.Enabled = TF;
                    textBox_Group_Count_1.Enabled = TF;
                    textBox_Group_People_1.Enabled = TF;
                    textBox_Group_Amount_1.Enabled = TF;
                    textBox_Lanwair1_1.Enabled = TF;
                    textBox_Lanwair2_1.Enabled = TF;
                    textBox_VavmTake_1.Enabled = TF;
                    textBox_Repair_1.Enabled = TF;

                    #endregion

                    #region Group Pao- 1

                    textBox_ConcessionRefund_1.Enabled = TF;
                    textBox_RefundNotice_1.Enabled = TF;
                    textBox_VavmRefund_1.Enabled = TF;
                    textBox_PvRefund_1.Enabled = TF;
                    textBox_CreditCardRefund_1.Enabled = TF;
                    textBox_QrCodeRefund_1.Enabled = TF;
                    textBox_SaleCardRefund_1.Enabled = TF;
                    textBox_GroupRefund_1.Enabled = TF;

                    #endregion

                    #region Group Exchange Pao 1

                    textBox_ExchangeTicket_20_1.Enabled = TF;
                    textBox_ExchangeTicket_25_1.Enabled = TF;
                    textBox_ExchangeTicket_30_1.Enabled = TF;

                    #endregion

                    #region Group OneDay

                    textBox_ExchangeOneDay_1.Enabled = TF;
                    textBox_KKDay.Enabled = TF;
                    textBox_Ntpc.Enabled = TF;
                    textBox_SplitTicker.Enabled = TF;

                    #endregion

                }
                else if (item == 2)
                {

                    #region Label

                    label_pao_2_1.Enabled = TF;
                    label_pao_2_2.Enabled = TF;
                    label_pao_2_3.Enabled = TF;
                    label_pao_2_4.Enabled = TF;
                    label_Ticket_15.Enabled = TF;
                    label_Ticket_30.Enabled = TF;
                    label_ExchangeTicket_30.Enabled = TF;
                    #endregion

                    #region Group Pao+ 2

                    textBox_OneDay_2.Enabled = TF;
                    textBox_OneDay_E_2.Enabled = TF;
                    textBox_Ticket_10_2.Enabled = TF;
                    textBox_Ticket_12_2.Enabled = TF;
                    textBox_Ticket_15_2.Enabled = TF;
                    textBox_Ticket_20_2.Enabled = TF;
                    textBox_Ticket_25_2.Enabled = TF;
                    textBox_Ticket_30_2.Enabled = TF;
                    textBox_Ticket_50_2.Enabled = TF;
                    textBox_OpenTransaction_2.Enabled = TF;
                    textBox_Tpass_2.Enabled = TF;
                    textBox_Group_Count_2.Enabled = TF;
                    textBox_Group_People_2.Enabled = TF;
                    textBox_Group_Amount_2.Enabled = TF;
                    textBox_Lanwair1_2.Enabled = TF;
                    textBox_Lanwair2_2.Enabled = TF;
                    textBox_VavmTake_2.Enabled = TF;
                    textBox_Repair_2.Enabled = TF;

                    #endregion

                    #region Group Pao- 2

                    textBox_ConcessionRefund_2.Enabled = TF;
                    textBox_RefundNotice_2.Enabled = TF;
                    textBox_VavmRefund_2.Enabled = TF;
                    textBox_PvRefund_2.Enabled = TF;
                    textBox_CreditCardRefund_2.Enabled = TF;
                    textBox_QrCodeRefund_2.Enabled = TF;
                    textBox_SaleCardRefund_2.Enabled = TF;
                    textBox_GroupRefund_2.Enabled = TF;

                    #endregion

                    #region Group Exchange Pao 2

                    textBox_ExchangeTicket_20_2.Enabled = TF;
                    textBox_ExchangeTicket_25_2.Enabled = TF;
                    textBox_ExchangeTicket_30_2.Enabled = TF;

                    #endregion

                    #region Group OneDay

                    textBox_ExchangeOneDay_2.Enabled = TF;

                    #endregion

                }
            });
        }

        private void Pao_Status(bool TF)
        {
            MyCommond.WriteLog(ThisReceive, $"PAO Text - {((TF) ? "開啟" : "關閉")}");
            this.MyCommond.InvokeIfRequired(this, () =>
            {
                if (Form_PaoText_UI_Status.Count > 0)
                {

                    #region Label

                    label_pao_1_1.Enabled = TF;
                    label_pao_1_2.Enabled = TF;
                    label_pao_1_3.Enabled = TF;
                    label_pao_1_4.Enabled = TF;

                    #endregion

                    #region Group Pao+ 1

                    if (Form_PaoText_UI_Status[0].Ticket_OneDay == "1") /*       */ textBox_OneDay_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Ticket_OneDay == "1") /*       */ textBox_OneDay_E_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Ticket10 == "1") /*            */ textBox_Ticket_10_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Ticket12 == "1") /*            */ textBox_Ticket_12_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Ticket15 == "1") /*            */ textBox_Ticket_15_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Ticket20 == "1") /*            */ textBox_Ticket_20_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Ticket25 == "1") /*            */ textBox_Ticket_25_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Ticket30 == "1") /*            */ textBox_Ticket_30_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].TicketBike == "1") /*          */ textBox_Ticket_50_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].OpenTrafficSetTicket == "1") /**/ textBox_OpenTransaction_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].TPASS == "1") /*               */ textBox_Tpass_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].GroupCount == "1") /*          */ textBox_Group_Count_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].GroupPeople == "1") /*         */ textBox_Group_People_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].GroupAmt == "1") /*            */ textBox_Group_Amount_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Lanwair1 == "1") /*            */ textBox_Lanwair1_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Lanwair2 == "1") /*            */ textBox_Lanwair2_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].AfcTakeOut == "1") /*          */ textBox_VavmTake_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].RepairTicket == "1") /*        */ textBox_Repair_1.Enabled = TF;

                    #endregion

                    #region Group Pao- 1

                    if (Form_PaoText_UI_Status[0].ConcessionRefund == "1") /**/ textBox_ConcessionRefund_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].RefundNotice == "1") /*    */ textBox_RefundNotice_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].VavmRefund == "1") /*      */ textBox_VavmRefund_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].PvRefund == "1") /*        */ textBox_PvRefund_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].CreditCardRefund == "1") /**/ textBox_CreditCardRefund_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].QrCodeRefund == "1") /*    */ textBox_QrCodeRefund_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].SaleCardRefund == "1") /*  */ textBox_SaleCardRefund_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].GroupRefund == "1") /*     */ textBox_GroupRefund_1.Enabled = TF;

                    #endregion

                    #region Group Exchange Pao 1

                    if (Form_PaoText_UI_Status[0].ExchangeTicket20 == "1") /**/ textBox_ExchangeTicket_20_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].ExchangeTicket25 == "1") /**/ textBox_ExchangeTicket_25_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].ExchangeTicket30 == "1") /**/ textBox_ExchangeTicket_30_1.Enabled = TF;

                    #endregion

                    #region Group OneDay

                    if (Form_PaoText_UI_Status[0].ExchangeOneDay == "1") /**/ textBox_ExchangeOneDay_1.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].KKDay == "1") /*         */ textBox_KKDay.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Ntpc == "1") /*          */ textBox_Ntpc.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].SplitTicker == "1") /*   */ textBox_SplitTicker.Enabled = TF;

                    #endregion

                }
                if (Form_PaoText_UI_Status.Count > 1)
                {

                    #region Label

                    label_pao_2_1.Enabled = TF;
                    label_pao_2_2.Enabled = TF;
                    label_pao_2_3.Enabled = TF;
                    label_pao_2_4.Enabled = TF;
                    label_Ticket_15.Enabled = TF;
                    label_Ticket_30.Enabled = TF;
                    label_ExchangeTicket_30.Enabled = TF;
                    #endregion

                    #region Group Pao+ 2

                    if (Form_PaoText_UI_Status[0].Ticket_OneDay == "1") /*       */ textBox_OneDay_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Ticket_OneDay == "1") /*       */ textBox_OneDay_E_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Ticket10 == "1") /*            */ textBox_Ticket_10_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Ticket12 == "1") /*            */ textBox_Ticket_12_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Ticket15 == "1") /*            */ textBox_Ticket_15_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Ticket20 == "1") /*            */ textBox_Ticket_20_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Ticket25 == "1") /*            */ textBox_Ticket_25_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Ticket30 == "1") /*            */ textBox_Ticket_30_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].TicketBike == "1") /*          */ textBox_Ticket_50_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].OpenTrafficSetTicket == "1") /**/ textBox_OpenTransaction_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].TPASS == "1") /*               */ textBox_Tpass_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].GroupCount == "1") /*          */ textBox_Group_Count_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].GroupPeople == "1") /*         */ textBox_Group_People_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].GroupAmt == "1") /*            */ textBox_Group_Amount_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Lanwair1 == "1") /*            */ textBox_Lanwair1_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].Lanwair2 == "1") /*      */ textBox_Lanwair2_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].AfcTakeOut == "1") /*          */ textBox_VavmTake_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].RepairTicket == "1") /*        */ textBox_Repair_2.Enabled = TF;

                    #endregion

                    #region Group Pao- 2

                    if (Form_PaoText_UI_Status[0].ConcessionRefund == "1") /**/ textBox_ConcessionRefund_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].RefundNotice == "1") /*    */ textBox_RefundNotice_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].VavmRefund == "1") /*      */ textBox_VavmRefund_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].PvRefund == "1") /*        */ textBox_PvRefund_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].CreditCardRefund == "1") /**/ textBox_CreditCardRefund_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].QrCodeRefund == "1") /*    */ textBox_QrCodeRefund_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].SaleCardRefund == "1") /*  */ textBox_SaleCardRefund_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].GroupRefund == "1") /*     */ textBox_GroupRefund_2.Enabled = TF;

                    #endregion

                    #region Group Exchange Pao 2

                    if (Form_PaoText_UI_Status[0].ExchangeTicket20 == "1") /**/ textBox_ExchangeTicket_20_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].ExchangeTicket25 == "1") /**/ textBox_ExchangeTicket_25_2.Enabled = TF;
                    if (Form_PaoText_UI_Status[0].ExchangeTicket30 == "1") /**/ textBox_ExchangeTicket_30_2.Enabled = TF;

                    #endregion

                    #region Group OneDay

                    if (Form_PaoText_UI_Status[0].ExchangeOneDay == "1") /**/ textBox_ExchangeOneDay_2.Enabled = TF;

                    #endregion

                }
            });
        }

        private bool StringToBool(string TF)
        {
            if (TF == "1") return true;
            return false;
        }

        private void Button_Status(bool TF)
        {
            MyCommond.WriteLog(ThisReceive, $"各報表按鈕 - {((TF) ? "開啟" : "關閉")}");
            MyCommond.InvokeIfRequired(this, () =>
            {
                if (MainServer.globalCommond.Finish_Button >= 1) button1.Enabled = TF;       //營運資料統計表(運量)
                if (MainServer.globalCommond.Finish_Button >= 2) button2.Enabled = TF;       //營運資料統計表(營收)
                if (MainServer.globalCommond.Finish_Button >= 3) button3.Enabled = TF;       //每日全線各站分時運量日報表
                if (MainServer.globalCommond.Finish_Button >= 4) button4.Enabled = TF;       //票卡進出站運量日報表
                //if (MainServer.globalCommonds.Finish_Button >= 5) button5.Enabled = TF;       //日運量統計表  /* 永遠上鎖 */
                if (MainServer.globalCommond.Finish_Button >= 6) button6.Enabled = TF;       //營收日報表
                if (MainServer.globalCommond.Finish_Button >= 7) button7.Enabled = TF;       //每日起訖站總表
                if (MainServer.globalCommond.Finish_Button >= 8) button8.Enabled = TF;       //設備營收日報表
                //if (MainServer.globalCommonds.Finish_Button >= 9) button9.Enabled = TF;       //月運量統計表 /* 永遠上鎖 */
                if (MainServer.globalCommond.Finish_Button >= 10) button10.Enabled = TF;     //每月起訖總表
                if (MainServer.globalCommond.Finish_Button >= 11) button11.Enabled = TF;     //各家票卡運量月報表(車站)
                if (MainServer.globalCommond.Finish_Button >= 12) button12.Enabled = TF;     //各家票卡運量月報表(天期)
                if (MainServer.globalCommond.Finish_Button >= 13) button13.Enabled = TF;     //自有票卡運量月報表(車站)
                if (MainServer.globalCommond.Finish_Button >= 14) button14.Enabled = TF;     //自有票卡運量月報表(天期)
                if (MainServer.globalCommond.Finish_Button >= 15) button15.Enabled = TF;     //營收月報表(車站)
                if (MainServer.globalCommond.Finish_Button >= 16) button16.Enabled = TF;     //營收月報表(天期)
                if (MainServer.globalCommond.Finish_Button >= 17) button17.Enabled = TF;     //設備營收月報表(車站)
                if (MainServer.globalCommond.Finish_Button >= 18) button18.Enabled = TF;     //設備營收月報表(天期)
                if (MainServer.globalCommond.Finish_Button >= 19) button19.Enabled = TF;     //營運資料月報(運量)
                if (MainServer.globalCommond.Finish_Button >= 20) button20.Enabled = TF;     //營運資料月報(營收)
                if (Test_Button) button_Test.Enabled = TF;  //Test Button
            });
        }

        private void Other_Status(bool TF)
        {
            MyCommond.WriteLog(ThisReceive, $"其他物件 - {((TF) ? "開啟" : "關閉")}");
            MyCommond.InvokeIfRequired(this, () =>
            {
                button_Test.Enabled = TF;
                textBox0.Enabled = TF;
                textBox_CreditCard.Enabled = TF;
                button_CreditCard.Enabled = TF;
                textBox_QRCode.Enabled = TF;
                button_QRCode.Enabled = TF;
                dateTimePicker1.Enabled = TF;
                //dateTimePicker2.Enabled = TF;
            });
        }

        #endregion

        private void ButtonMethode(LrtTxnClientSwitch Methode)
        {
            All_Item_Status(false);

            MyCommond.WriteLog(ThisReceive, $"準備取得PAO資料");
            Form_PaoText_Data = Form2Data();
            System.Threading.Thread threadStartTimer = new System.Threading.Thread(Timer_Start);
            threadStartTimer.IsBackground = true;
            if (MyCheckBoxs.FindIndex(x => x.Status == true) < 0) goto NoStart;
            int SearchStatus = (int)MyCheckBoxs.Find(x => x.Status == true).Methode;

            if ((int)Methode > 0) checkIO(Methode, true);
            if (SearchStatus < 9)
            {
                MyCommond.WriteLog(ThisReceive, $"確認是日報表");
                if ((int)Methode > 0)
                {
                    MyCommond.WriteLog(ThisReceive, $"有指定報表");
                    MainServer.OperationStartDate = Convert.ToDateTime(dateTimePicker2.Value.ToString("yyyy/MM/dd 05:00:00"));
                    MainServer.OperationEndDate = Convert.ToDateTime(dateTimePicker1.Value.ToString("yyyy/MM/dd 05:00:00"));
                }
                else
                {
                    MyCommond.WriteLog(ThisReceive, $"沒指定報表");
                    MainServer.OperationStartDate = Convert.ToDateTime(dateTimePicker3.Value.ToString("yyyy/MM/dd 05:00:00"));
                    MainServer.OperationEndDate = Convert.ToDateTime(dateTimePicker4.Value.ToString("yyyy/MM/dd 05:00:00")).AddDays(1);
                }
                if (!CheckMessage(Form_PaoText_Data)) goto NoStart;

                ///CheckAndExportData(new ExportConfig()
                ///{
                ///    Path1 = MainServer.Path_PaoData,
                ///    DTime = new DateTime(),
                ///    FirstName = dateTimePicker2.Value.ToString("yyyyMM"),
                ///    FileName = MainServer.FileName_PaoCsv,
                ///    SubName = "",
                ///    OpLine = MainServer.OperationLine_String,
                ///    LastName = "csv",
                ///}, new PaoTicketSaleList().ExportTitle_zhTW(), Form_PaoText_Data, Encoding.UTF8);

                var s = new ExportConfig()
                {
                    Path1 = MainServer.Path_PaoData,
                    DTime = new DateTime(),
                    FirstName = dateTimePicker2.Value.ToString("yyyyMM"),
                    FileName = MainServer.FileName_PaoCsv,
                    SubName = "",
                    OpLine = MainServer.globalCommond.OperationCode_String,
                    LastName = "csv",
                };
                MyCommond.ExportStuctData(ThisReceive, s, new PaoTicketSaleList().ExportTitle_zhTW(), Form_PaoText_Data, Encoding.UTF8);
                /// if (NeedExport)
                /// {
                ///     var s = new ExportConfig()
                ///     {
                ///         Path1 = MainServer.Path_PaoData,
                ///         DTime = new DateTime(),
                ///         FirstName = dateTimePicker2.Value.ToString("yyyyMM"),
                ///         FileName = MainServer.FileName_PaoCsv,
                ///         SubName = "",
                ///         OpLine = MainServer.globalCommond.OperationLine_String,
                ///         LastName = "csv",
                ///     };
                ///     MyCommond.ExportStuctData(ThisReceive, s, new PaoTicketSaleList().ExportTitle_zhTW(), Form_PaoText_Data, Encoding.UTF8);
                /// }
            }
            else
            {
                MyCommond.WriteLog(ThisReceive, $"確認是月報表");
                if ((int)Methode > 0)
                {
                    MyCommond.WriteLog(ThisReceive, $"有指定報表");
                    MainServer.OperationStartDate = Convert.ToDateTime(dateTimePicker2.Value.ToString("yyyy/MM/01 05:00:00"));
                    MainServer.OperationEndDate = MainServer.OperationStartDate.AddMonths(1);
                }
                else
                {
                    MyCommond.WriteLog(ThisReceive, $"沒指定報表");
                    MainServer.OperationStartDate = Convert.ToDateTime(dateTimePicker5.Value.ToString("yyyy/MM/dd 05:00:00"));
                    MainServer.OperationEndDate = Convert.ToDateTime(dateTimePicker6.Value.ToString("yyyy/MM/dd 05:00:00")).AddDays(1);
                }
            }

            MyCommond.WriteLog(ThisReceive, $"準備取得資料庫");
            System.Threading.Thread threadStartMain = new System.Threading.Thread(() =>
            {
                MainServer.Start(MyCheckBoxs, Form_PaoText_Data);
            });
            threadStartMain.IsBackground = true;
            threadStartMain.Start();
            threadStartTimer.Start();
            threadStartMain.Join();
            goto EndSub;
        NoStart:
            All_Item_Status(true);
        EndSub:
            Console.WriteLine("End");
            MyCommond.WriteLog(ThisReceive, $"End");
        }

        #endregion

        #region Event

        #region UI Event - Normal

        #endregion

        #region UI Event - progressBar

        #endregion

        #region UI Event - Text

        #region Group Pao+

        #region Pao 1

        private void textBox_OneDay_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_OneDay_1.Text = CheckInPut(ref e, textBox_OneDay_1.Text);
        }

        private void textBox_OneDay_E_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_OneDay_E_1.Text = CheckInPut(ref e, textBox_OneDay_E_1.Text);
        }

        private void textBox_Ticket_10_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Ticket_10_1.Text = CheckInPut(ref e, textBox_Ticket_10_1.Text);
        }

        private void textBox_Ticket_12_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Ticket_12_1.Text = CheckInPut(ref e, textBox_Ticket_12_1.Text);
        }

        private void textBox_Ticket_15_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Ticket_15_1.Text = CheckInPut(ref e, textBox_Ticket_15_1.Text);
        }

        private void textBox_Ticket_20_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Ticket_20_1.Text = CheckInPut(ref e, textBox_Ticket_20_1.Text);
        }

        private void textBox_Ticket_25_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Ticket_25_1.Text = CheckInPut(ref e, textBox_Ticket_25_1.Text);
        }

        private void textBox_Ticket_30_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Ticket_30_1.Text = CheckInPut(ref e, textBox_Ticket_30_1.Text);
        }

        private void textBox_Ticket_50_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Ticket_50_1.Text = CheckInPut(ref e, textBox_Ticket_50_1.Text);
        }

        private void textBox_OpenTransaction_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_OpenTransaction_1.Text = CheckInPut(ref e, textBox_OpenTransaction_1.Text);
        }

        private void textBox_Tpass_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Tpass_1.Text = CheckInPut(ref e, textBox_Tpass_1.Text);
        }

        private void textBox_Group_Count_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Group_Count_1.Text = CheckInPut(ref e, textBox_Group_Count_1.Text);
        }

        private void textBox_Group_People_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Group_People_1.Text = CheckInPut(ref e, textBox_Group_People_1.Text);
        }

        private void textBox_Group_Amount_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Group_Amount_1.Text = CheckInPut(ref e, textBox_Group_Amount_1.Text);
        }

        private void textBox_Lanwair1_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Lanwair1_1.Text = CheckInPut(ref e, textBox_Lanwair1_1.Text);
        }

        private void textBox_Lanwair2_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Lanwair2_1.Text = CheckInPut(ref e, textBox_Lanwair2_1.Text);
        }

        private void textBox_VavmTake_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_VavmTake_1.Text = CheckInPut(ref e, textBox_VavmTake_1.Text);
        }

        private void textBox_Repair_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Repair_1.Text = CheckInPut(ref e, textBox_Repair_1.Text);
        }

        #endregion

        #region Pao 2

        private void textBox_OneDay_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_OneDay_2.Text = CheckInPut(ref e, textBox_OneDay_2.Text);
        }

        private void textBox_OneDay_E_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_OneDay_E_2.Text = CheckInPut(ref e, textBox_OneDay_E_2.Text);
        }

        private void textBox_Ticket_10_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Ticket_10_2.Text = CheckInPut(ref e, textBox_Ticket_10_2.Text);
        }

        private void textBox_Ticket_12_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Ticket_12_2.Text = CheckInPut(ref e, textBox_Ticket_12_2.Text);
        }

        private void textBox_Ticket_15_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Ticket_15_2.Text = CheckInPut(ref e, textBox_Ticket_15_2.Text);
        }

        private void textBox_Ticket_20_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Ticket_20_2.Text = CheckInPut(ref e, textBox_Ticket_20_2.Text);
        }

        private void textBox_Ticket_25_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Ticket_25_2.Text = CheckInPut(ref e, textBox_Ticket_25_2.Text);
        }

        private void textBox_Ticket_30_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Ticket_30_2.Text = CheckInPut(ref e, textBox_Ticket_30_2.Text);
        }

        private void textBox_Ticket_50_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Ticket_50_2.Text = CheckInPut(ref e, textBox_Ticket_50_2.Text);
        }

        private void textBox_OpenTransaction_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_OpenTransaction_2.Text = CheckInPut(ref e, textBox_OpenTransaction_2.Text);
        }

        private void textBox_Tpass_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Tpass_2.Text = CheckInPut(ref e, textBox_Tpass_2.Text);
        }

        private void textBox_Group_Count_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Group_Count_2.Text = CheckInPut(ref e, textBox_Group_Count_2.Text);
        }

        private void textBox_Group_People_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Group_People_2.Text = CheckInPut(ref e, textBox_Group_People_2.Text);
        }

        private void textBox_Group_Amount_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Group_Amount_2.Text = CheckInPut(ref e, textBox_Group_Amount_2.Text);
        }

        private void textBox_Lanwair1_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Lanwair1_2.Text = CheckInPut(ref e, textBox_Lanwair1_2.Text);
        }

        private void textBox_Lanwair2_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Lanwair2_2.Text = CheckInPut(ref e, textBox_Lanwair2_2.Text);
        }

        private void textBox_VavmTake_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_VavmTake_2.Text = CheckInPut(ref e, textBox_VavmTake_2.Text);
        }

        private void textBox_Repair_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Repair_2.Text = CheckInPut(ref e, textBox_Repair_2.Text);
        }

        #endregion

        #endregion

        #region Group Pao-

        #region Pao 1

        private void textBox_ConcessionRefund_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_ConcessionRefund_1.Text = CheckInPut(ref e, textBox_ConcessionRefund_1.Text);
        }

        private void textBox_RefundNotice_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_RefundNotice_1.Text = CheckInPut(ref e, textBox_RefundNotice_1.Text);
        }

        private void textBox_VavmRefund_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_VavmRefund_1.Text = CheckInPut(ref e, textBox_VavmRefund_1.Text);
        }

        private void textBox_PvRefund_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_PvRefund_1.Text = CheckInPut(ref e, textBox_PvRefund_1.Text);
        }

        private void textBox_CreditCardRefund_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_CreditCardRefund_1.Text = CheckInPut(ref e, textBox_CreditCardRefund_1.Text);
        }

        private void textBox_QrCodeRefund_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_QrCodeRefund_1.Text = CheckInPut(ref e, textBox_QrCodeRefund_1.Text);
        }

        private void textBox_SaleCardRefund_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_SaleCardRefund_1.Text = CheckInPut(ref e, textBox_SaleCardRefund_1.Text);
        }

        private void textBox_GroupRefund_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_GroupRefund_1.Text = CheckInPut(ref e, textBox_GroupRefund_1.Text);
        }

        #endregion

        #region Pao 2

        private void textBox_ConcessionRefund_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_ConcessionRefund_2.Text = CheckInPut(ref e, textBox_ConcessionRefund_2.Text);
        }

        private void textBox_RefundNotice_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_RefundNotice_2.Text = CheckInPut(ref e, textBox_RefundNotice_2.Text);
        }

        private void textBox_VavmRefund_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_VavmRefund_2.Text = CheckInPut(ref e, textBox_VavmRefund_2.Text);
        }

        private void textBox_PvRefund_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_PvRefund_2.Text = CheckInPut(ref e, textBox_PvRefund_2.Text);
        }

        private void textBox_CreditCardRefund_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_CreditCardRefund_2.Text = CheckInPut(ref e, textBox_CreditCardRefund_2.Text);
        }

        private void textBox_QrCodeRefund_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_QrCodeRefund_2.Text = CheckInPut(ref e, textBox_QrCodeRefund_2.Text);
        }

        private void textBox_SaleCardRefund_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_SaleCardRefund_2.Text = CheckInPut(ref e, textBox_SaleCardRefund_2.Text);
        }

        private void textBox_GroupRefund_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_GroupRefund_2.Text = CheckInPut(ref e, textBox_GroupRefund_2.Text);
        }

        #endregion

        #endregion

        #region Group OneDay

        private void textBox_ExchangeOneDay_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_ExchangeOneDay_1.Text = CheckInPut(ref e, textBox_ExchangeOneDay_1.Text);
        }

        private void textBox_ExchangeOneDay_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_ExchangeOneDay_2.Text = CheckInPut(ref e, textBox_ExchangeOneDay_2.Text);
        }

        private void textBox_KKDay_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_KKDay.Text = CheckInPut(ref e, textBox_KKDay.Text);
        }

        private void textBox_Ntpc_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_Ntpc.Text = CheckInPut(ref e, textBox_Ntpc.Text);
        }

        private void textBox_SplitTicker_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_SplitTicker.Text = CheckInPut(ref e, textBox_SplitTicker.Text);
        }

        #endregion

        #region Group Exchange

        #region Pao 1

        private void textBox_ExchangeTicket_20_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_ExchangeTicket_20_1.Text = CheckInPut(ref e, textBox_ExchangeTicket_20_1.Text);
        }

        private void textBox_ExchangeTicket_25_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_ExchangeTicket_25_1.Text = CheckInPut(ref e, textBox_ExchangeTicket_25_1.Text);
        }

        private void textBox_ExchangeTicket_30_1_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_ExchangeTicket_30_1.Text = CheckInPut(ref e, textBox_ExchangeTicket_30_1.Text);
        }

        #endregion

        #region Pao 2

        private void textBox_ExchangeTicket_20_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_ExchangeTicket_20_2.Text = CheckInPut(ref e, textBox_ExchangeTicket_20_2.Text);
        }

        private void textBox_ExchangeTicket_25_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_ExchangeTicket_25_2.Text = CheckInPut(ref e, textBox_ExchangeTicket_25_2.Text);
        }

        private void textBox_ExchangeTicket_30_2_Changed(object sender, KeyPressEventArgs e)
        {
            textBox_ExchangeTicket_30_2.Text = CheckInPut(ref e, textBox_ExchangeTicket_30_2.Text);
        }

        #endregion

        #endregion

        #region Other

        private void TextBox0_Changed(object sender, KeyPressEventArgs e)
        {
            //textBox0.Text = CheckInPut(ref e, textBox0.Text);
        }

        #endregion

        private string CheckInPut(ref KeyPressEventArgs e, string Text)
        {
            string reStr = Text;
            if (((int)e.KeyChar < 48 | (int)e.KeyChar > 57) & (int)e.KeyChar != 8)
            {
                e.Handled = true;
            }
            if (Text.Trim() != "")
            {
                if (Convert.ToInt32(Text.Trim()) <= 0)
                {
                    reStr = "0";
                }
            }
            else
            {
                //textBox3.Text = "0";
            }
            return reStr;
        }

        #endregion

        #region UI Event - CheckBox

        #region Group1

        //是否使用線程
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"啟用線程 - {((checkBox1.Checked) ? "啟動" : "關閉")}");
            MainServer.Thread_Use = checkBox1.Checked;
            if (ConfigLoading)
            {
                //WriteLog(0, MethodBase.GetCurrentMethod().Name.ToString(), "初始化中暫不處裡");
            }
            else
            {
                //WriteLog(0, MethodBase.GetCurrentMethod().Name.ToString(), "線程啟用狀態更新：" + Convert.ToString(Thread_Use));
            }

        }

        //是否顯示Excel
        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"開啟背景程式 - {((checkBox2.Checked) ? "啟動" : "關閉")}");
            MainServer.globalCommond.Excel_Applaction_Show = checkBox2.Checked;
            if (ConfigLoading)
            {
                //WriteLog(0, MethodBase.GetCurrentMethod().Name.ToString(), "初始化中暫不處裡");
            }
            else
            {
                //WriteLog(0, MethodBase.GetCurrentMethod().Name.ToString(), "Excel顯示啟用狀態更新：" + Convert.ToString(Excel_Applaction_Show));
            }

        }

        //是否自動帶入班結表
        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"帶入班傑表 - {((checkBox3.Checked) ? "啟動" : "關閉")}");
            MainServer.globalCommond.Auto_Pao_Excel = checkBox3.Checked;
            if (ConfigLoading)
            {
                //WriteLog(0, MethodBase.GetCurrentMethod().Name.ToString(), "初始化中暫不處裡");
            }
            else
            {
                //WriteLog(0, MethodBase.GetCurrentMethod().Name.ToString(), "帶入班結表啟用狀態更新：" + Convert.ToString(Auto_Pao_Excel));
            }

        }

        //是否載入行動支付
        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"帶入電子支付 - {((checkBox4.Checked) ? "啟動" : "關閉")}");
            MainServer.globalCommond.Install_MobilePay = checkBox4.Checked;
            if (ConfigLoading)
            {
                //WriteLog(0, MethodBase.GetCurrentMethod().Name.ToString(), "初始化中暫不處裡");
            }
            else
            {
                //WriteLog(0, MethodBase.GetCurrentMethod().Name.ToString(), "行動支付載入狀態更新：" + Convert.ToString(Install_MobilePay));
            }

        }

        #endregion

        #region Group2

        private void checkBox_G_1_1_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_1_1); }

        private void checkBox_G_1_2_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_1_2); }

        private void checkBox_G_1_3_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_1_3); }

        private void checkBox_G_1_4_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_1_4); }

        private void checkBox_G_1_5_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_1_5); }

        private void checkBox_G_1_6_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_1_6); }

        private void checkBox_G_1_7_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_1_7); }

        private void checkBox_G_1_8_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_1_8); }

        #endregion

        #region Group3

        private void checkBox_G_2_1_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_2_1); }

        private void checkBox_G_2_2_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_2_2); }

        private void checkBox_G_2_3_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_2_3); }

        private void checkBox_G_2_4_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_2_4); }

        private void checkBox_G_2_5_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_2_5); }

        private void checkBox_G_2_6_CheckedChanged(object sender, EventArgs e)  { if (!checkBatch) checkIO(checkBox_G_2_6); }

        private void checkBox_G_2_7_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_2_7); }

        private void checkBox_G_2_8_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_2_8); }

        private void checkBox_G_2_9_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_2_9); }

        private void checkBox_G_2_10_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_2_10); }

        private void checkBox_G_2_11_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_2_11); }

        private void checkBox_G_2_12_CheckedChanged(object sender, EventArgs e) { if (!checkBatch) checkIO(checkBox_G_2_12); }

        #endregion

        #endregion

        #region UI Event - DateTimePicker

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (ConfigLoading)
            {
                MyCommond.WriteLog(ThisReceive, $"初始化中暫不處裡");
                //WriteLog(0, MethodBase.GetCurrentMethod().Name.ToString(), "初始化中暫不處裡");
            }
            else
            {
                ConfigLoading = true;
                MainServer.OperationEndDate = Convert.ToDateTime(dateTimePicker1.Value.ToString("yyyy/MM/dd 05:00:00"));
                MainServer.OperationStartDate = MainServer.OperationEndDate.AddDays(-1);

                dateTimePicker1.Value = MainServer.OperationEndDate;
                MyCommond.WriteLog(ThisReceive, $"更新產製日期:{Convert.ToString(MainServer.OperationEndDate)}");
                //WriteLog(0, MethodBase.GetCurrentMethod().Name.ToString(), "更新產製日期:" + Convert.ToString(Operation_End_Date));
                dateTimePicker2.Value = MainServer.OperationStartDate;
                dateTimePicker3.Value = MainServer.OperationStartDate;
                //dateTimePicker3.Value = Operation_Start_Date;


                MyCommond.WriteLog(ThisReceive, $"寫入行動支付路徑及檔案");
                //WriteLog(0, MethodBase.GetCurrentMethod().Name.ToString(), "寫入行動支付路徑及檔案");
                textBox_CreditCard.Text = MyCommond.Path_Program + MyCommond.Path_Analyze + $@"信用卡每日交易表-" + MainServer.OperationEndDate.ToString("yyyy-MM-dd") + ".csv";
                textBox_QRCode.Text = MyCommond.Path_Program + MyCommond.Path_Analyze + $@"QRCode每日交易表-" + MainServer.OperationStartDate.ToString("yyyy-MM-dd") + ".csv";
                if (MainServer.globalCommond.Auto_Pao_Excel)
                {
                    CheckPaoCsv = false;
                    if (MainServer.globalCommond.Auto_Pao_Excel)
                    {
                        //All_Item_Status(false);
                        //PaoSaleList();
                        //Data2Form(Form_PaoText_Data);
                        //All_Item_Status(true);
                        System.Threading.Thread thread = new System.Threading.Thread(() =>
                        {
                            All_Item_Status(false);
                            GetPaoSaleList();
                            Data2Form(Form_PaoText_Data);
                            All_Item_Status(true);
                        });
                        thread.IsBackground = true;
                        thread.Start();
                    }
                }
                else
                {
                    CheckPaoCsv = false;
                    MyCommond.WriteLog(ThisReceive, $"未啟用自動帶入班結表");
                    //WriteLog(0, MethodBase.GetCurrentMethod().Name.ToString(), "未啟用自動帶入班結表");
                }
                ConfigLoading = false;
            }
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"更新營運日期:{Convert.ToString(dateTimePicker2.Value)}");
            //WriteLog(0, MethodBase.GetCurrentMethod().Name.ToString(), "更新營運日期:" + Convert.ToString(Operation_Start_Date));
        }

        //批量時間變化
        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"更新營運日期:{Convert.ToString(dateTimePicker3.Value)}");
            //WriteLog(0, MethodBase.GetCurrentMethod().Name.ToString(), "更新開始日期:" + Convert.ToString(dateTimePicker3.Value));
            dateTimePicker4.Value = dateTimePicker3.Value; ;
        }

        private void dateTimePicker4_ValueChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"更新營運日期:{Convert.ToString(dateTimePicker4.Value)}");
            //WriteLog(0, MethodBase.GetCurrentMethod().Name.ToString(), "更新結束日期:" + Convert.ToString(dateTimePicker4.Value));
        }

        private void dateTimePicker5_ValueChanged(object sender, EventArgs e)
        {
            dateTimePicker5.Value = Convert.ToDateTime(dateTimePicker5.Value.ToString("yyyy/MM/dd 05:00:00"));
            MyCommond.WriteLog(ThisReceive, $"更新營運日期:{Convert.ToString(dateTimePicker5.Value)}");
            //WriteLog(0, MethodBase.GetCurrentMethod().Name.ToString(), "更新開始日期:" + Convert.ToString(dateTimePicker5.Value));
            DateTime mDate = Convert.ToDateTime(dateTimePicker5.Value.AddMonths(1).ToString("yyyy/MM/01 05:00:00"));
            dateTimePicker6.Value = mDate.AddDays(-1);
        }

        private void dateTimePicker6_ValueChanged(object sender, EventArgs e)
        {
            MyCommond.WriteLog(ThisReceive, $"更新營運日期:{Convert.ToString(dateTimePicker6.Value)}");
            //WriteLog(0, MethodBase.GetCurrentMethod().Name.ToString(), "更新結束日期:" + Convert.ToString(dateTimePicker6.Value));
        }

        #endregion

        #region UI Event - Button

        #region Group1

        private void Button1_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Day_Volume).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Day_Volume);
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Day_Amount).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Day_Amount);
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Day_EachStation_EachTime).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Day_EachStation_EachTime);
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Day_ElectronicTicket).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Day_ElectronicTicket);
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Day_AllRideList).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Day_AllRideList);
        }

        private void Button6_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Day_TrafficAmount).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Day_TrafficAmount);
        }

        private void Button7_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Day_OriginDestination).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Day_OriginDestination);
        }

        private void Button8_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Day_EquipAmount).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Day_EquipAmount);
        }

        #endregion

        #region Group2

        private void Button9_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Month_AllRideList).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Month_AllRideList);
        }

        private void Button10_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Month_OriginDestination).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Month_OriginDestination);
        }

        private void Button11_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Month_ElectronicTicket_Station).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Month_ElectronicTicket_Station);
        }

        private void Button12_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Month_ElectronicTicket_Day).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Month_ElectronicTicket_Day);
        }

        private void Button13_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Month_OwnTicketVolume_Station).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Month_OwnTicketVolume_Station);
        }

        private void Button14_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Month_OwnTicketVolume_Day).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Month_OwnTicketVolume_Day);
        }

        private void Button15_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Month_TrafficAmount_Station).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Month_TrafficAmount_Station);
        }

        private void Button16_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Month_TrafficAmount_Day).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Month_TrafficAmount_Day);
        }

        private void Button17_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Month_EquipAmount_Station).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Month_EquipAmount_Station);
        }

        private void Button18_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Month_EquipAmount_Day).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Month_EquipAmount_Day);
        }

        private void Button19_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Month_Volume).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Month_Volume);
        }

        private void Button20_Click(object sender, EventArgs e)
        {
            MyCheckBoxs.Find(x => x.Methode == LrtTxnClientSwitch.Month_Amount).Status = true;
            ButtonMethode(LrtTxnClientSwitch.Month_Amount);
        }

        #endregion

        #region Group3

        private void Button_G_1_Start_Click(object sender, EventArgs e)
        {
            ButtonMethode(LrtTxnClientSwitch.NA);
        }

        private void Button_G_1_All_Click(object sender, EventArgs e)
        {
            checkIO(DatePark.Day, true);
        }

        private void Button_G_1_Cancel_Click(object sender, EventArgs e)
        {
            checkIO(DatePark.Day, false);
        }

        #endregion

        #region Group4

        private void Button_G_2_Start_Click(object sender, EventArgs e)
        {
            ButtonMethode(LrtTxnClientSwitch.NA);
        }

        private void Button_G_2_All_Click(object sender, EventArgs e)
        {
            checkIO(DatePark.Month, true);
        }

        private void Button_G_2_Cancel_Click(object sender, EventArgs e)
        {
            checkIO(DatePark.Month, false);
        }

        #endregion

        #region Other

        private void Button_CreditCard_Click(object sender, EventArgs e)
        {
            string ss = MyCommond.SelectFileForm(new FileLastName() { csv = true });
            if (ss != "")
            {
                MyCommond.InvokeIfRequired(this, () =>
                {
                    textBox_CreditCard.Text = ss;
                    toolTip1.SetToolTip(textBox_CreditCard, ss);
                });
            }
        }

        private void Button_QRCode_Click(object sender, EventArgs e)
        {
            string ss = MyCommond.SelectFileForm(new FileLastName() { csv = true });
            if (ss != "")
            {
                MyCommond.InvokeIfRequired(this, () =>
                {
                    textBox_QRCode.Text = ss;
                    toolTip1.SetToolTip(textBox_QRCode, ss);
                });
            }
        }

        private void Button_Test_Click(object sender, EventArgs e)
        {
            All_Item_Status(false);
            System.Threading.Thread thread = new System.Threading.Thread(() =>
            {
                string sql_command_string = @"Data Source='{0}';Initial Catalog={1};Persist Security Info=True;User ID='{2}';Password='{3}'; Timeout=60 ";
                string sql_command = string.Format(sql_command_string, MainServer.globalCommond.SqlServers.IP, MainServer.globalCommond.SqlServers.Calalog, MainServer.globalCommond.SqlServers.ID, MainServer.globalCommond.SqlServers.Password);
                var Station = (from p in MainServer.globalCommond.Stock_Lrt_Station select Convert.ToByte(p.CodeNumber)).ToList();
                string Sqlstr = @"--
DECLARE @StartStation  VARCHAR(10)  --該路線車站起點
DECLARE @EndStation    VARCHAR(10)  --該路線車站終點
DECLARE @Openstardate  DATETIME     --搜尋時間開始
DECLARE @Openenddate   DATETIME     --搜尋時間結束
DECLARE @Monthstardate DATETIME     --開始的月份
DECLARE @Monthenddate  DATETIME     --結束的月份
DECLARE @DateSet       DATETIME     --計算月份用
DECLARE @RunDate       DATETIME     --執行的月份
DECLARE @in_columns    VARCHAR(100) --存放車站別

SET @StartStation  = '{0}' 
SET @EndStation    = '{1}' 
SET @Openstardate  = '{2}' 
SET @Openenddate   = '{3}' 
SET @DateSet       = CONCAT(DATEPART(YEAR,@Openstardate),'-',DATEPART(MONTH,@Openstardate) , '-01 0' , DATEPART(HOUR,@Openstardate) , ':0' , DATEPART(MINUTE,@Openstardate) , ':0' , DATEPART(SECOND,@Openstardate))
SET @Monthenddate  = DATEADD(MONTH,1,@DateSet)
SET @Monthstardate = DATEADD(MONTH,-1,@Monthenddate)
SET @RunDate       = @Monthstardate

SELECT LocationId AS mLoc INTO #mTempLocation FROM Parm081_LocList WHERE LocationId >= @StartStation AND LocationId <= @EndStation AND CurrentType IN (1)

SELECT 
    CASE 
    WHEN TName IN ('進站','出站') THEN 
        CASE 
        WHEN TxnType  IN (29,30) THEN 
            CASE 
            WHEN DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND ((IssueCode IN (2) AND PersonalProfile IN (5)) OR (IssueCode IN (9) AND CardType IN (7) AND AreaCode IN (21,22,23,25,30,31,32,33,35,36,37,39,40,41,43,44,45,47,48,49,51,52,53,55,56,57,59,60,61,63,64,65,67,68,69,71,72,73,75,76,77,79,80,81,83,84,85,87,88,89,91,92,93,95,96,97,99,100,101,103,104,105,107,108,109,111,112,113))) 
                THEN '行政院通勤月票_學生卡' 
            ELSE '行政院通勤月票_一般卡'
            END 
        WHEN TxnType  IN (21,22) THEN 
            CASE
            WHEN DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND ((IssueCode IN (2) AND PersonalProfile IN (5)) OR (IssueCode IN (9) AND CardType IN (7) AND AreaCode IN (21,22,23,25,30,31,32,33,35,36,37,39,40,41,43,44,45,47,48,49,51,52,53,55,56,57,59,60,61,63,64,65,67,68,69,71,72,73,75,76,77,79,80,81,83,84,85,87,88,89,91,92,93,95,96,97,99,100,101,103,104,105,107,108,109,111,112,113))) 
                THEN '學生卡'
            WHEN DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND ((IssueCode = 2 AND PersonalProfile IN (8) AND AreaCode IN (1, 2)) OR (IssueCode = 9 AND CardType IN (7) AND AreaCode IN (78,82))) THEN '兒童卡'
            WHEN DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND ((IssueCode = 2 AND PersonalProfile IN (1, 2)                    ) OR (IssueCode = 9 AND CardType IN (1)                        )) THEN '敬老卡'
            WHEN DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND ((IssueCode = 2 AND PersonalProfile IN (3)                       ) OR (IssueCode = 9 AND CardType IN (2)                        )) THEN '愛心卡'
            WHEN DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND ((IssueCode = 2 AND PersonalProfile IN (4)                       ) OR (IssueCode = 9 AND CardType IN (3)                        )) THEN 
                CASE
                WHEN WelfareAmt IN (0) THEN '一般卡'
                ELSE '陪伴卡'
                END
            ELSE '一般卡'
            END
        WHEN TxnType IN (23,24) THEN
            CASE
            WHEN IssueCode  = 11 AND PeriodCode    = 16 AND OriginalTimes = 20 THEN '海樂地'
            WHEN PeriodCode = 17 AND ValidPeriod   = 40 /*------------------*/ THEN '員工票'
            WHEN PeriodCode = 1  AND ValidPeriod   = 1  /*------------------*/ THEN '一日票'
            WHEN PeriodCode = 16 AND OriginalTimes = 1  /*------------------*/ THEN '退小月票'
            WHEN PeriodCode = 1  AND ValidPeriod   = 31 /*------------------*/ THEN '小月票'
            ELSE '效期票'
            END
        END
    WHEN TName IN ('售票','退票') THEN
        CASE
        WHEN CardType IN (1)     THEN '一般票'
        WHEN CardType IN (3)     THEN '優待票'
        WHEN CardType IN (2)     THEN '自行車票'
        ELSE '其他票'
        END
    END AS TicketType
    ,*
INTO #TempAllTxn
FROM
(
          SELECT '進站'     AS TName,    LocationId,         EquipType,                 CardSN,    TxnType,    TxnDT,    IssueCode,         CardType,         PersonalProfile,         IdentityType,         IdentityExpiryDT,    TxnAmt,         StartLoc,         EndLoc,  NULL AS SaleStartLoc,  NULL AS SaleEndLoc,         TxnWP,         TxnPoint, NULL AS TrnAmt,         AreaCode, NULL AS WelfareAmt,         PeriodCode,         OriginalTimes,         ValidPeriod, NULL AS BodyData, NULL AS LastTrnFlag FROM Txn_Entry               JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate 
    UNION SELECT '出站'     AS TName, T1.LocationId,      T1.EquipType,              T1.CardSN, T1.TxnType, T1.TxnDT, T1.IssueCode,      T1.CardType,      T1.PersonalProfile,      T1.IdentityType,      T1.IdentityExpiryDT, T1.TxnAmt,      T1.StartLoc,      T1.EndLoc,  NULL AS SaleStartLoc,  NULL AS SaleEndLoc,      T1.TxnWP,      T1.TxnPoint,      T1.TrnAmt,      T1.AreaCode,      T1.WelfareAmt,      T1.PeriodCode,      T1.OriginalTimes,      T1.ValidPeriod,      T2.BodyData,      T1.LastTrnFlag FROM (SELECT * FROM Txn_Exit JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate) AS T1 LEFT JOIN (SELECT * FROM ECCTxn_Interface WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate) AS T2 ON T1.CardSN = T2.CardSN AND T1.TxnDT = T2.TxnDT AND T1.TxnType = T2.MyTxnType
    UNION SELECT '售票'     AS TName,    LocationId,         EquipType, TicketUnique AS CardSN,    TxnType,    TxnDT,    IssueCode,         CardType, NULL AS PersonalProfile, NULL AS IdentityType, NULL AS IdentityExpiryDT,    TxnAmt,         StartLoc,         EndLoc,  NULL AS SaleStartLoc,  NULL AS SaleEndLoc, NULL AS TxnWP, NULL AS TxnPoint, NULL AS TrnAmt, NULL AS AreaCode, NULL AS WelfareAmt, NULL AS PeriodCode, NULL AS OriginalTimes, NULL AS ValidPeriod, NULL AS BodyData, NULL AS LastTrnFlag FROM Txn_SellSpecialTicket   JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate 
    UNION SELECT '退票'     AS TName,    LocationId,         EquipType, TicketUnique AS CardSN,    TxnType,    TxnDT,    IssueCode,         CardType, NULL AS PersonalProfile, NULL AS IdentityType, NULL AS IdentityExpiryDT,    TxnAmt, NULL AS StartLoc, NULL AS EndLoc,          SaleStartLoc,          SaleEndLoc, NULL AS TxnWP, NULL AS TxnPoint, NULL AS TrnAmt, NULL AS AreaCode, NULL AS WelfareAmt, NULL AS PeriodCode, NULL AS OriginalTimes, NULL AS ValidPeriod, NULL AS BodyData, NULL AS LastTrnFlag FROM Txn_RefundSpeTkt        JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate 
    UNION SELECT '罰款'     AS TName,    LocationId,         EquipType,                 CardSN,    TxnType,    TxnDT,    IssueCode,         CardType,         PersonalProfile,         IdentityType,         IdentityExpiryDT,    TxnAmt, NULL AS StartLoc, NULL AS EndLoc,  NULL AS SaleStartLoc,  NULL AS SaleEndLoc,         TxnWP, NULL AS TxnPoint, NULL AS TrnAmt, NULL AS AreaCode, NULL AS WelfareAmt, NULL AS PeriodCode, NULL AS OriginalTimes, NULL AS ValidPeriod, NULL AS BodyData, NULL AS LastTrnFlag FROM Txn_ExcessByCSC         JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate 
    UNION SELECT 'PAM補票'  AS TName,    LocationId,         EquipType,                 CardSN,    TxnType,    TxnDT,    IssueCode,         CardType,         PersonalProfile,         IdentityType,         IdentityExpiryDT,    TxnAmt,         StartLoc,         EndLoc,  NULL AS SaleStartLoc,  NULL AS SaleEndLoc, NULL AS TxnWP,         TxnPoint,         TrnAmt,         AreaCode,         WelfareAmt,         PeriodCode,         OriginalTimes,         ValidPeriod, NULL AS BodyData,         LastTrnFlag FROM Txn_ExceptDeduct        JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate 
    UNION SELECT '售卡'     AS TName,    LocationId,         EquipType,                 CardSN,    TxnType,    TxnDT,    IssueCode,         CardType, NULL AS PersonalProfile, NULL AS IdentityType, NULL AS IdentityExpiryDT,    TxnAmt, NULL AS StartLoc, NULL AS EndLoc,  NULL AS SaleStartLoc,  NULL AS SaleEndLoc, NULL AS TxnWP, NULL AS TxnPoint, NULL AS TrnAmt, NULL AS AreaCode, NULL AS WelfareAmt, NULL AS PeriodCode, NULL AS OriginalTimes, NULL AS ValidPeriod, NULL AS BodyData, NULL AS LastTrnFlag FROM Txn_SaleCard            JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate 
    UNION SELECT '加值'     AS TName,    LocationId,         EquipType,                 CardSN,    TxnType,    TxnDT,    IssueCode,         CardType,         PersonalProfile,         IdentityType,         IdentityExpiryDT,    TxnAmt, NULL AS StartLoc, NULL AS EndLoc,  NULL AS SaleStartLoc,  NULL AS SaleEndLoc, NULL AS TxnWP, NULL AS TxnPoint, NULL AS TrnAmt, NULL AS AreaCode, NULL AS WelfareAmt, NULL AS PeriodCode, NULL AS OriginalTimes, NULL AS ValidPeriod, NULL AS BodyData, NULL AS LastTrnFlag FROM Txn_AddValue            JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate 
    UNION SELECT '取消加值' AS TName,    LocationId,         EquipType,                 CardSN,    TxnType,    TxnDT,    IssueCode,         CardType,         PersonalProfile,         IdentityType,         IdentityExpiryDT,    TxnAmt, NULL AS StartLoc, NULL AS EndLoc,  NULL AS SaleStartLoc,  NULL AS SaleEndLoc, NULL AS TxnWP, NULL AS TxnPoint, NULL AS TrnAmt, NULL AS AreaCode, NULL AS WelfareAmt, NULL AS PeriodCode, NULL AS OriginalTimes, NULL AS ValidPeriod, NULL AS BodyData, NULL AS LastTrnFlag FROM Txn_CancelValue         JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate 
    UNION SELECT '自動加值' AS TName,    LocationId, NULL AS EquipType,                 CardSN,    TxnType,    TxnDT,    IssueCode, NULL AS CardType, NULL AS PersonalProfile, NULL AS IdentityType, NULL AS IdentityExpiryDT,    TxnAmt, NULL AS StartLoc, NULL AS EndLoc,  NULL AS SaleStartLoc,  NULL AS SaleEndLoc, NULL AS TxnWP, NULL AS TxnPoint, NULL AS TrnAmt, NULL AS AreaCode, NULL AS WelfareAmt, NULL AS PeriodCode, NULL AS OriginalTimes, NULL AS ValidPeriod, NULL AS BodyData, NULL AS LastTrnFlag FROM ECCTxn_Interface        JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType = 2  AND MyTxnType = 2 
    UNION SELECT '自動加值' AS TName,    LocationId, NULL AS EquipType,                 CardSN,    TxnType,    TxnDT,    IssueCode,         CardType, NULL AS PersonalProfile, NULL AS IdentityType, NULL AS IdentityExpiryDT,    TxnAmt, NULL AS StartLoc, NULL AS EndLoc,  NULL AS SaleStartLoc,  NULL AS SaleEndLoc, NULL AS TxnWP, NULL AS TxnPoint, NULL AS TrnAmt, NULL AS AreaCode, NULL AS WelfareAmt, NULL AS PeriodCode, NULL AS OriginalTimes, NULL AS ValidPeriod, NULL AS BodyData, NULL AS LastTrnFlag FROM iPASSTxn_OATI           JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType = 66 
    UNION SELECT '自動加值' AS TName,    LocationId, NULL AS EquipType,                 CardSN,    TxnType,    TxnDT,    IssueCode,         CardType, NULL AS PersonalProfile, NULL AS IdentityType, NULL AS IdentityExpiryDT,    TxnAmt, NULL AS StartLoc, NULL AS EndLoc,  NULL AS SaleStartLoc,  NULL AS SaleEndLoc, NULL AS TxnWP,         TxnPoint, NULL AS TrnAmt, NULL AS AreaCode, NULL AS WelfareAmt, NULL AS PeriodCode, NULL AS OriginalTimes, NULL AS ValidPeriod, NULL AS BodyData, NULL AS LastTrnFlag FROM icasHTxn_ICTX           JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType = 2  AND TxnSubType = 4 
) AS aa

SELECT *
FROM #TempAllTxn

DROP TABLE #TempAllTxn
DROP TABLE #mTempLocation
";
                string commandText = string.Format(Sqlstr, Station.Min(), Station.Max(), dateTimePicker2.Value.ToString("yyyy/MM/dd 05:00:00"), dateTimePicker1.Value.ToString("yyyy/MM/dd 05:00:00")); ;
                DataTable dt = new DataTable();
                int Reconnect = 0;

                string strHostName = string.Empty;
                System.Net.IPHostEntry ipEntry = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName());
                System.Net.IPAddress[] addr = ipEntry.AddressList;

                for (int i = 0; i < addr.Length; i++)
                {
                    MyCommond.WriteLog(ThisReceive, $"IP Address {i}: {addr[i].MapToIPv4().ToString()}");

                    string[] sss = addr[i].MapToIPv4().ToString().Split('.');
                    string[] yyy = MainServer.globalCommond.SqlServers.IP.Split('.');
                    int cc = 0;

                    for (int j = 0; j < yyy.Length; j++)
                    {
                        if (sss[j] == yyy[j]) { cc++; }
                    }

                    if (cc == 3) { goto StartConnect0; }
                }
                return;

            StartConnect0:
                MyCommond.WriteLog(ThisReceive, $"準備資料庫連線");
            StartConnect:
                using (System.Data.SqlClient.SqlConnection con = new System.Data.SqlClient.SqlConnection(sql_command))
                {
                    using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand())
                    {
                        cmd.CommandText = commandText;
                        cmd.CommandType = CommandType.Text;
                        cmd.Connection = con;
                        MyCommond.WriteLog(ThisReceive, $"寫入資料庫語法\n{commandText}");

                        try
                        {
                            con.Open();
                            System.Data.SqlClient.SqlDataReader reader = cmd.ExecuteReader();

                            MyCommond.WriteLog(ThisReceive, $"載入資料庫取得的資料");
                            dt.Load(reader);
                        }
                        catch (Exception ex)
                        {
                            if (ex.Message == "執行逾時到期。在作業完成之前超過逾時等待的時間，或是伺服器未回應。")
                            {
                                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
                            }
                            else
                            {
                                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{MyCommond.ExportErrorMessageToLog(ex, "全交易資料表").Message}");
                            }
                        }
                        finally
                        {
                            con.Close();
                        }
                    }
                }

                if (Reconnect >= 3) return;

                if (dt == null || dt.Rows.Count <= 0)
                {
                    MyCommond.WriteLog(ThisReceive, $"重新查詢");
                    Reconnect++;
                    goto StartConnect;
                }

                var dr = dt.CreateDataReader();
                var dr_c = dr.FieldCount;
                //string[] ResponTitle = new string[dr_c];
                string mTitle = "";
                for (int i = 0; i < dr_c; i++)
                {
                    MyCommond.WriteLog(ThisReceive, $"dr.GetName({i}) = {dr.GetName(i)}");
                    //ResponTitle[i] = dr.GetName(i);
                    mTitle += dr.GetName(i) + ",";
                }

                MyCommond.WriteLog(ThisReceive, $"將資料轉換成數字陣列");
                MyCommond.WriteLog(ThisReceive, $"行:{dt.Rows.Count,3}, 欄:{dt.Columns.Count,3}");
                string[,] vs = new string[dt.Rows.Count, dt.Columns.Count];
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        vs[i, j] = Convert.ToString(dt.Rows[i][j]);
                    }
                }

                MyCommond.ExportData(ThisReceive, new ExportConfig()
                {
                    Path1 = MyCommond.Path_Program,
                    DTime = dateTimePicker1.Value,
                    FileName = "全交易資料明細",
                    LastName = "xlsx"
                }, mTitle, vs, Encoding.UTF8);

                System.Threading.Thread thread1 = new System.Threading.Thread(() =>
                {
                    MessageBox.Show("完成資料匯出");
                });
                thread1.IsBackground = true;
                thread1.Start();

                All_Item_Status(true);
            });
            thread.IsBackground = true;
            thread.Start();
        }

        #endregion

        #endregion

        #endregion

        #endregion

        private bool CheckMessage(List<PaoTicketSaleList> PaoData_UI)
        {
            MyCommond.WriteLog(ThisReceive, $"二次確認無營收一日票");
            bool RePort = false;

            string FullShowText = $"\n" +
                $"KKDay\t\t：{PaoData_UI[0].KKDay,3} 張\n" +
                $"新北幣\t\t：{PaoData_UI[0].Ntpc,3} 張\n" +
                $"切票\t\t：{PaoData_UI[0].SplitTicker,3} 張\n" +
                $"一日票兌換({PaoData_UI[0].Station})\t：{PaoData_UI[0].ExchangeOneDay,3} 張\n" +
                $"單程票20元({PaoData_UI[0].Station})\t：{PaoData_UI[0].ExchangeTicket20,3} 張\n" +
                $"單程票25元({PaoData_UI[0].Station})\t：{PaoData_UI[0].ExchangeTicket25,3} 張\n" +
                $"單程票30元({PaoData_UI[0].Station})\t：{PaoData_UI[0].ExchangeTicket30,3} 張\n";

            if (PaoData_UI.Count > 1)
            {
                FullShowText += $"===================\n" +
                    $"一日票兌換({PaoData_UI[1].Station})\t：{PaoData_UI[1].ExchangeOneDay,3} 張\n" +
                    $"單程票20元({PaoData_UI[1].Station})\t：{PaoData_UI[1].ExchangeTicket20,3} 張\n" +
                    $"單程票25元({PaoData_UI[1].Station})\t：{PaoData_UI[1].ExchangeTicket25,3} 張\n" +
                    $"單程票30元({PaoData_UI[1].Station})\t：{PaoData_UI[1].ExchangeTicket30,3} 張\n";
            }

            FullShowText += "是否正確?";

            MyCommond.WriteLog(ThisReceive, FullShowText);
            var CheckMessageBox = MessageBox.Show(FullShowText, "再次確認", MessageBoxButtons.YesNo);
            MyCommond.WriteLog(ThisReceive, $"取得使用者回覆：{CheckMessageBox.ToString()}");
            if (CheckMessageBox == DialogResult.Yes) RePort = true;
            else RePort = false;
            MyCommond.WriteLog(ThisReceive, $"Return RePort Staute {RePort}");
            return RePort;
        }

    }
}
