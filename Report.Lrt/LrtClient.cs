using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//-
using System.IO;
using System.Threading;
using System.Windows.Forms;
using UsuallyCommond.MyEnum;
using ClosedXML.Excel;
using System.Data;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
//using Report.Lrt.ViewModel;
//using Report.Lrt.Load;
using GlobalCommond;
using GlobalCommond.ViewModel;
using System.Net;

namespace Report.Lrt
{
    public class LrtClient
    {
        #region 設定

        public Guid ID { get; set; }
        public LrtMain MainServer { get; set; }
        public LrtTxnClientSwitch MethodeName { get; set; }
        public StatusEnum ThisMethode { get; set; }
        public DateTime OperationDate { get; set; }
        public int MaxCount { get; set; }
        public int MinCount { get; set; }
        public int ValueCount { get; set; }
        #endregion

        #region 變數

        private MyCommonds MyCommond;
        private string ThisReceive;
        private Thread threadStart;

        private string ExportPath;
        private string ExportFileName;

        #endregion

        #region 資料
        public List<MobilePay> QrCodeData { get; set; }  //QRCode每日交易表-2022-10-15.csv
        public List<MobilePay> CreditCardData { get; set; }  //信用卡每日交易表-2022-11-06.csv

        #endregion

        public LrtClient()
        {
            MyCommond = new MyCommonds();
            ThisMethode = StatusEnum.Disconnected;
            MaxCount = 100;
            MinCount = 0;
            ValueCount = 0;
        }

        public LrtClient(LrtMain GetMainServer, LrtTxnClientSwitch mMothendName) : this()
        {
            MainServer = GetMainServer;
            MethodeName = mMothendName;
            ID = Guid.NewGuid();
            ThisReceive = $@"client\{MethodeName.ToString()}({ID.ToString()})";
            ThisMethode = StatusEnum.Ready;
            MyCommond.WriteLog(ThisReceive, "創建新Clien端");
        }

        public void Test()
        {
            string[] item0 = null;
            var item1 = DataToList(GetSqlData(""), ref item0);
            string item2 = "";
            for (int i = 0; i < item0.Length; i++) { item2 += item0[i]; }
            MyCommond.ExportData(ThisReceive, new ExportConfig()
            {
                Path1 = MyCommond.Path_Program,
                FileName = "item",
                LastName = "csv",
                FirstName = OperationDate.ToString("yyyyMMdd"),
            }, item2, item1, Encoding.UTF8);
            ThisMethode = StatusEnum.Ready;
        }

        #region 功能

        public void Start()
        {
            threadStart = new Thread(SendingMethod);
            threadStart.IsBackground = true;
            threadStart.Start();
        }

        #endregion

        #region 非同步線呈執行程式

        int DateCount = 0;
        private void SendingMethod(object obj)
        {
            MyCommond.WriteLog(ThisReceive, $"準備計算 - {MethodeName.ToDescription()}");
            ValueCount = 10;
            string LastFileName = "xlsx";

            string MethodeReportName = MainServer.globalCommond.Stock_ReportFile.Find(x => x.OperationLine == MainServer.globalCommond.OperationCode && x.ReportNameEng == MethodeName.ToString()).ReportNameEng;
            string ExportFullPath = $"{MyCommond.Path_Program + MyCommond.Path_Template}{MethodeReportName}.{LastFileName}";

            while (ThisMethode != StatusEnum.Disconnected)
            {
                if (ThisMethode == StatusEnum.Ready)
                {
                         if (MethodeName >= LrtTxnClientSwitch.Month_AllRideList)
                    {

                        string StationOrDay = "";
                        int Int_Row = 0;
                        int Int_Column = 30;

                        switch (MethodeName)
                        {
                            case LrtTxnClientSwitch.Month_AllRideList /*              */ : StationOrDay = "NA"; /*------*/ break;
                            case LrtTxnClientSwitch.Month_OriginDestination /*        */ : StationOrDay = "Station"; /*-*/ break;
                            case LrtTxnClientSwitch.Month_ElectronicTicket_Station /* */ : StationOrDay = "Station"; /*-*/ break;
                            case LrtTxnClientSwitch.Month_ElectronicTicket_Day /*     */ : StationOrDay = "Day"; /*-----*/ break;
                            case LrtTxnClientSwitch.Month_OwnTicketVolume_Station /*  */ : StationOrDay = "Station"; /*-*/ break;
                            case LrtTxnClientSwitch.Month_OwnTicketVolume_Day /*      */ : StationOrDay = "Day"; /*-----*/ break;
                            case LrtTxnClientSwitch.Month_TrafficAmount_Station /*    */ : StationOrDay = "Station"; /*-*/ break;
                            case LrtTxnClientSwitch.Month_TrafficAmount_Day /*        */ : StationOrDay = "Day"; /*-----*/ break;
                            case LrtTxnClientSwitch.Month_EquipAmount_Station /*      */ : StationOrDay = "Station"; /*-*/ break;
                            case LrtTxnClientSwitch.Month_EquipAmount_Day /*          */ : StationOrDay = "Day"; /*-----*/ break;
                            case LrtTxnClientSwitch.Month_Volume /*                   */ : StationOrDay = "Day"; /*-----*/ break;
                            case LrtTxnClientSwitch.Month_Amount /*                   */ : StationOrDay = "Day"; /*-----*/ break;
                            default: goto ExportMonthEnd;
                        }

                        if (StationOrDay == "Station")  { Int_Row = MainServer.globalCommond.Stock_Lrt_Station.Count; }
                        else if (StationOrDay == "Day") { Int_Row = Convert.ToDateTime(MainServer.OperationStartDate.ToString("yyyy/MM/01")).AddMonths(1).AddDays(-1).Day; }

                        int[,] Month_Data_Afc = new int[Int_Row, Int_Column];
                        int[,] Month_Data_MobilePay = new int[Int_Row, Int_Column];
                        int[,] Month_Data_OwnTicket = new int[Int_Row, Int_Column];

                        OperationDate = Convert.ToDateTime(MainServer.OperationStartDate.ToString("yyyy/MM/01 05:00:00"));
                        ReportMessage();
                        if (!DataLoad_Month())
                        {
                            MyCommond.WriteLog(ThisReceive, "結束執行。");

                            goto ExportMonthEnd;
                        }
                        FileNameExample(new ExportConfig()
                        {
                            Path1 = MyCommond.Path_Program + MyCommond.Path_Report,
                            DTime = MainServer.OperationStartDate,
                            FirstName = "",
                            FileName = MainServer.globalCommond.Stock_ReportFile.Find(x => x.OperationLine == MainServer.globalCommond.OperationCode && x.ReportNameEng == MethodeName.ToString()).ReportNameCh,
                            SubName = "",
                            OpLine = MainServer.globalCommond.OperationCode_String,
                            LastName = "xlsx",
                        });

                        ValueCount = 20;
                        switch (MethodeName)
                        {
                            case LrtTxnClientSwitch.Month_AllRideList /*              */ : Calculator_Month_AllRideList /*              */ (ref Month_Data_Afc, ref Month_Data_MobilePay, ref Month_Data_OwnTicket); break;
                            case LrtTxnClientSwitch.Month_OriginDestination /*        */ : Calculator_Month_OriginDestination /*        */ (ref Month_Data_Afc, ref Month_Data_MobilePay, ref Month_Data_OwnTicket); break;
                            case LrtTxnClientSwitch.Month_ElectronicTicket_Station /* */ : Calculator_Month_ElectronicTicket_Station /* */ (ref Month_Data_Afc, ref Month_Data_MobilePay, ref Month_Data_OwnTicket); break;
                            case LrtTxnClientSwitch.Month_ElectronicTicket_Day /*     */ : Calculator_Month_ElectronicTicket_Day /*     */ (ref Month_Data_Afc, ref Month_Data_MobilePay, ref Month_Data_OwnTicket); break;
                            case LrtTxnClientSwitch.Month_OwnTicketVolume_Station /*  */ : Calculator_Month_OwnTicketVolume_Station /*  */ (ref Month_Data_Afc, ref Month_Data_MobilePay, ref Month_Data_OwnTicket); break;
                            case LrtTxnClientSwitch.Month_OwnTicketVolume_Day /*      */ : Calculator_Month_OwnTicketVolume_Day /*      */ (ref Month_Data_Afc, ref Month_Data_MobilePay, ref Month_Data_OwnTicket); break;
                            case LrtTxnClientSwitch.Month_TrafficAmount_Station /*    */ : Calculator_Month_TrafficAmount_Station /*    */ (ref Month_Data_Afc, ref Month_Data_MobilePay, ref Month_Data_OwnTicket); break;
                            case LrtTxnClientSwitch.Month_TrafficAmount_Day /*        */ : Calculator_Month_TrafficAmount_Day /*        */ (ref Month_Data_Afc, ref Month_Data_MobilePay, ref Month_Data_OwnTicket); break;
                            case LrtTxnClientSwitch.Month_EquipAmount_Station /*      */ : Calculator_Month_EquipAmount_Station /*      */ (ref Month_Data_Afc, ref Month_Data_MobilePay, ref Month_Data_OwnTicket); break;
                            case LrtTxnClientSwitch.Month_EquipAmount_Day /*          */ : Calculator_Month_EquipAmount_Day /*          */ (ref Month_Data_Afc, ref Month_Data_MobilePay, ref Month_Data_OwnTicket); break;
                            case LrtTxnClientSwitch.Month_Volume /*                   */ : Calculator_Month_Volume /*                   */ (ref Month_Data_Afc, ref Month_Data_MobilePay, ref Month_Data_OwnTicket); break;
                            case LrtTxnClientSwitch.Month_Amount /*                   */ : Calculator_Month_Amount /*                   */ (ref Month_Data_Afc, ref Month_Data_MobilePay, ref Month_Data_OwnTicket); break;
                            default: goto ExportMonthEnd;
                        }
                        ValueCount = 80;

                        MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Month_Data_Afc, "Day_Data_Afc");
                        MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Month_Data_MobilePay, "Day_Data_MobilePay");
                        MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Month_Data_OwnTicket, "Day_Data_OwnTicket");

                        ExportExcle_V1(ExportFullPath, Month_Data_Afc, Month_Data_MobilePay, Month_Data_OwnTicket);

                    ExportMonthEnd:
                        ValueCount = 95;
                    } /* 月報 */
                    else if (MethodeName <  LrtTxnClientSwitch.Month_AllRideList)
                    {
                        DateCount = (MainServer.OperationEndDate - MainServer.OperationStartDate).Days;
                        for (OperationDate = MainServer.OperationStartDate; OperationDate < MainServer.OperationEndDate; OperationDate = OperationDate.AddDays(1))
                        {
                            int[,] Day_Data_Afc = null;
                            int[,] Day_Data_MobilePay = null;
                            int[,] Day_Data_OwnTicket = null;

                            DataLoad_Day();
                            ReportMessage();
                            FileNameExample(new ExportConfig()
                            {
                                Path1 = MyCommond.Path_Program + MyCommond.Path_Report,
                                DTime = OperationDate,
                                FirstName = "",
                                //FileName = MainServer.Stock_Lrt_ReportFile.Find(x => x.ReportNameEng == MethodeName.ToString()).ReportNameCh,
                                FileName = MainServer.globalCommond.Stock_ReportFile.Find(x => x.OperationLine == MainServer.globalCommond.OperationCode && x.ReportNameEng == MethodeName.ToString()).ReportNameCh,
                                SubName = "",
                                OpLine = MainServer.globalCommond.OperationCode_String,
                                LastName = "xlsx",
                            });

                            ValueCount = 20;
                            switch (MethodeName)
                            {
                                case LrtTxnClientSwitch.NA /*                      */ : Test(); /*                             */ goto ExportDayEnd;
                                case LrtTxnClientSwitch.Day_Volume /*              */ : Calculator_Day_Volume /*               */ (ref Day_Data_Afc, ref Day_Data_MobilePay, ref Day_Data_OwnTicket); ExportFullPath = MyCommond.Path_Report; break;
                                case LrtTxnClientSwitch.Day_Amount /*              */ : Calculator_Day_Amount /*               */ (ref Day_Data_Afc, ref Day_Data_MobilePay, ref Day_Data_OwnTicket); ExportFullPath = MyCommond.Path_Report; break;
                                case LrtTxnClientSwitch.Day_EachStation_EachTime /**/ : Calculator_Day_EachStation_EachTime /* */ (ref Day_Data_Afc, ref Day_Data_MobilePay, ref Day_Data_OwnTicket); break;
                                case LrtTxnClientSwitch.Day_ElectronicTicket /*    */ : Calculator_Day_ElectronicTicket /*     */ (ref Day_Data_Afc, ref Day_Data_MobilePay, ref Day_Data_OwnTicket); break;
                                case LrtTxnClientSwitch.Day_AllRideList /*         */ : Calculator_Day_AllRideList /*          */ (ref Day_Data_Afc, ref Day_Data_MobilePay, ref Day_Data_OwnTicket); break;
                                case LrtTxnClientSwitch.Day_TrafficAmount /*       */ : Calculator_Day_TrafficAmount /*        */ (ref Day_Data_Afc, ref Day_Data_MobilePay, ref Day_Data_OwnTicket); break;
                                case LrtTxnClientSwitch.Day_OriginDestination /*   */ : Calculator_Day_OriginDestination /*    */ (ref Day_Data_Afc, ref Day_Data_MobilePay, ref Day_Data_OwnTicket); break;
                                case LrtTxnClientSwitch.Day_EquipAmount /*         */ : Calculator_Day_EquipAmount /*          */ (ref Day_Data_Afc, ref Day_Data_MobilePay, ref Day_Data_OwnTicket); break;
                                default: goto ExportDayEnd;
                            }
                            ValueCount = 80;

                            MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Day_Data_Afc, "Day_Data_Afc");
                            MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Day_Data_MobilePay, "Day_Data_MobilePay");
                            MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Day_Data_OwnTicket, "Day_Data_OwnTicket");

                            ExportExcle_V2(ExportFullPath, Day_Data_Afc, Day_Data_MobilePay, Day_Data_OwnTicket);

                            if (Day_MobilePayData2 == null || !(Day_MobilePayData2.Count > 0))
                            {
                                goto ExportDayEnd;
                            }
                            else
                            {
                                if (MainServer.ClientLock == "")
                                {
                                    MainServer.ClientLock = "a";
                                    MyCommond.ExportData(ThisReceive, new ExportConfig
                                    {
                                        Path1 = MyCommond.Path_Program + MyCommond.Path_Report,
                                        DTime = OperationDate,
                                        FirstName = "",
                                        FileName = "行動支付交易",
                                        SubName = "剔退",
                                        OpLine = MainServer.globalCommond.OperationCode_String,
                                        LastName = "xlsx",
                                    }, new MobilePay().ExportTitle_zhTW(), Day_MobilePayData2, Encoding.UTF8);
                                    Day_MobilePayData2 = null;
                                }
                            }

                        ExportDayEnd:
                            ThisMethode = StatusEnum.Next;
                            ValueCount = 95;
                        }
                    } /* 日報 */
                    else
                    {

                    }
                    
                    ValueCount += 1;

                    ThisMethode = StatusEnum.Disconnected;
                }
                Thread.Sleep(30);
            }
            ValueCount = 100;
            ThisMethode = StatusEnum.Finish;
        }

        private void ReportMessage()
        {
            MyCommond.WriteLog(ThisReceive, $"*****回報全域變數值*****");
            MyCommond.WriteLog(ThisReceive, $"開始日期:{MainServer.OperationStartDate}");
            MyCommond.WriteLog(ThisReceive, $"結束日期:{MainServer.OperationEndDate}");
            MyCommond.WriteLog(ThisReceive, $"營運日期:{OperationDate}");
        }

        #region 計算

        #region 日

        List<MobilePay> Day_MobilePayData; //僅載入有運量的
        List<MobilePay> Day_MobilePayData2; //儲存沒有運量的，僅偵錯
        List<GroupTicketList> Day_GroupTicket;

        private void DataLoad_Day()
        {
            Day_MobilePayData = new List<MobilePay>();
            Day_MobilePayData2 = new List<MobilePay>();
            Day_GroupTicket = new List<GroupTicketList>();

            #region 電支

            string MobilePayPath;
            var Station_StationName = (from p in MainServer.globalCommond.Stock_Lrt_Station select p.StationName).ToList();
            for (int i = 0; i < 2; i++)
            {
                if (i == 0)
                {
                    MobilePayPath = MyCommond.Path_Program + MyCommond.Path_Analyze + String.Format(MainServer.FileName_CreditCard, OperationDate.ToString("yyyy-MM-dd"));
                }
                else if (i == 1)
                {
                    MobilePayPath = MyCommond.Path_Program + MyCommond.Path_Analyze + String.Format(MainServer.FileName_QrCode, OperationDate.ToString("yyyy-MM-dd"));
                }
                else
                {
                    goto LoadNext_1;
                }
                MyCommond.WriteLog(ThisReceive, $"本次撈取電支路徑:{MobilePayPath}");
                ValueCount += 5;

                var I_Temp = MyCommond.ToViewModel<MobilePay>(ThisReceive, MobilePayPath, (new MobilePay()).ExportTitle_zhTW(), 2);
                MyCommond.WriteLog(ThisReceive, $"以撈取資料數量:{I_Temp.Count}");
                foreach (var item in I_Temp)
                {
                    MyCommond.WriteLog(ThisReceive, $"將分配的資料:{item.ToString()}");
                    if (
                            item.Volume == MobilePayVolum.InTime
                            || !(Station_StationName.Contains(item.EntryStation)
                            || Station_StationName.Contains(item.ExitStation))
                            || !((item.ExitTime >= OperationDate && item.ExitTime <= OperationDate.AddDays(1)) || item.ExitTime.Year == 1970) // 1970 之後要取消
                        )  // 剔退的條件
                    {
                        MyCommond.WriteLog(ThisReceive, $"將資料寫入剔退");
                        Day_MobilePayData2.Add(item);
                    }
                    else
                    {
                        MyCommond.WriteLog(ThisReceive, $"將資料寫入計算");
                        Day_MobilePayData.Add(item);
                    }
                }
                ValueCount += 5;
            }

            ListDataCount(Day_MobilePayData);

        #endregion

        LoadNext_1:
            #region 團體票 (只能算運量)

            var Ride_Date = OperationDate.ToString("yyyyMMdd");
            string GroupTicketPath = $"{MyCommond.Path_Program}{MainServer.Path_PaoData}" +
            $"{Ride_Date}_{MainServer.FileName_GroupTicket}_{MainServer.globalCommond.OperationCode_String}.csv";
            MyCommond.WriteLog(ThisReceive, $"本次讀取團體票路徑:{GroupTicketPath}");

            if (MyCommond.CheckFile(GroupTicketPath))
            {
                List<GroupTicketList> item = MyCommond.ToViewModel<GroupTicketList>(ThisReceive, GroupTicketPath, new GroupTicketList().ExportTitle_zhTW(), 1);

                foreach (var jtem in item)
                {
                    MyCommond.WriteLog(ThisReceive, $"讀取的資料:{jtem.ToString()}");
                    if (jtem.SaleDate != Ride_Date)
                    {
                        MyCommond.WriteLog(ThisReceive, $"非購買日使用，將使用時間鎖定至使用日上午十點");
                        jtem.SaleTime = "10:00";
                    }
                }

                Day_GroupTicket.AddRange(item);
            }

            ValueCount += 5;

        #endregion

        LoadNext_2:
            var thisover = 0;
        }

        private void MethodeIsRun(LrtTxnClientSwitch SearchMethode)
        {
            bool WaitRun = true;
            while (WaitRun)
            {
                Thread.Sleep(1000);
                var SearchClient = MainServer.LrtClients.Find(x => x.MethodeName == SearchMethode);
                var SearchClientCount = MainServer.LrtClients.FindIndex(x => x.MethodeName == SearchMethode);
                if (SearchClientCount > -1)
                {
                                 WaitRun = !(SearchClient.OperationDate == OperationDate);
                    if (WaitRun) WaitRun = !(SearchClient.OperationDate.AddDays(1) < OperationDate);
                }
                else WaitRun = false;
            }
        }

        #region 營運資料統計_運量

        private void Calculator_Day_Volume_Old(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            MethodeIsRun(LrtTxnClientSwitch.Day_Amount);
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                var SqlFileName = MainServer.globalCommond.Stock_SqlList.Find(x => x.UseNameCh == LrtTxnClientSwitch.Day_Volume.ToDescription()).FileName1;
                var Sqlstr = File.ReadAllText(MyCommond.Path_Program + MyCreat.Path_Sql + SqlFileName + ".txt");
                var Station_CodeNumber = (from p in MainServer.globalCommond.Stock_Lrt_Station select Convert.ToByte(p.CodeNumber)).ToList();
                var MobilePay_CC = (from p in MainServer.globalCommond.Stock_MobileCardType where p.CardType == "Credit" select p.CardSubType).ToList();
                var MobilePay_QQ = (from p in MainServer.globalCommond.Stock_MobileCardType where p.CardType == "QrCode" select p.CardSubType).ToList();
                var SqlCommand = string.Format(Sqlstr, Station_CodeNumber.Min(), Station_CodeNumber.Max(), OperationDate.ToString("yyyy/MM/dd 05:00:00"), OperationDate.AddDays(1).ToString("yyyy/MM/dd 05:00:00"));

                #region 人工
                MyCommond.WriteLog(ThisReceive, $"取得Pao收入、支出資料");
                var itemPao = MainServer.paoTicketSaleLists.FindAll(x => x.Operation == OperationDate.ToString("yyyyMMdd")).ToList();
                if (itemPao.Count == 0) { MyCommond.WriteLog(ThisReceive, "PaoData 異常"); }
                ValueCount += 1;

                MyCommond.WriteLog(ThisReceive, $"以團體票檔案載入");
                int GroupSumPeople = (from p in Day_GroupTicket select Convert.ToInt32(p.People)).ToList().Sum();
                ValueCount += 1;

                #endregion

                #region 電支
                MyCommond.WriteLog(ThisReceive, $"取得電子支付交易資料");
                var itemCridet = (from p in Day_MobilePayData where MobilePay_CC.Contains(p.CardSubType) select p).ToList();
                var itemQrCode = (from p in Day_MobilePayData where MobilePay_QQ.Contains(p.CardSubType) select p).ToList();
                ValueCount += 1;

                #endregion

                #region 電票
                MyCommond.WriteLog(ThisReceive, $"取得電子票證資料");
                string[] mTitle = new DayVolume().ExportTitle_zhTW().Split(',');
                string[,] SqlData = inputZ(new string[1, new DayVolume().ExportTitle_zhTW().Split(',').Length]);
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); SqlData = SQLTextSelect(SqlCommand, ref mTitle); }
                MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), SqlData, "SqlData");

                ValueCount += 1;

                #endregion

                MyCommond.WriteLog(ThisReceive, $"撰寫輸出資料");
                DayVolume ExportData = new DayVolume()
                {
                    Electron_Exit_num = Convert.ToString(Convert.ToInt32(SqlData[0, 0])),
                    Student_Exit_num = Convert.ToString(Convert.ToInt32(SqlData[0, 1])),
                    Welfare_Exit_num = Convert.ToString(Convert.ToInt32(SqlData[0, 2])),
                    All_Pass_Common_Exit_num = Convert.ToString(Convert.ToInt32(SqlData[0, 3])),
                    All_Pass_Student_Exit_num = Convert.ToString(Convert.ToInt32(SqlData[0, 4])),
                    Credit_num = Convert.ToString(itemCridet.Count()),
                    Qrcode_num = Convert.ToString(itemQrCode.Count()),
                    /*******************************************************************************************************************************/

                    SOneTkt_num = Convert.ToString(Convert.ToInt32(SqlData[0, 7]) +
                        ((itemPao.Count == 0) ? 0 : (itemPao.Count > 1) ?
                            Convert.ToInt32(itemPao[1].Ticket20) + Convert.ToInt32(itemPao[1].Ticket25) + Convert.ToInt32(itemPao[1].Ticket30) + Convert.ToInt32(itemPao[1].ExchangeTicket20) + Convert.ToInt32(itemPao[1].ExchangeTicket25) + Convert.ToInt32(itemPao[1].ExchangeTicket30) +
                            Convert.ToInt32(itemPao[0].Ticket20) + Convert.ToInt32(itemPao[0].Ticket25) + Convert.ToInt32(itemPao[0].Ticket30) + Convert.ToInt32(itemPao[0].ExchangeTicket20) + Convert.ToInt32(itemPao[0].ExchangeTicket25) + Convert.ToInt32(itemPao[0].ExchangeTicket30) :
                            Convert.ToInt32(itemPao[0].Ticket20) + Convert.ToInt32(itemPao[0].Ticket25) + Convert.ToInt32(itemPao[0].Ticket30) + Convert.ToInt32(itemPao[0].ExchangeTicket20) + Convert.ToInt32(itemPao[0].ExchangeTicket25) + Convert.ToInt32(itemPao[0].ExchangeTicket30))),
                    /*******************************************************************************************************************************/

                    SConcessionTkt_num = Convert.ToString(Convert.ToInt32(SqlData[0, 8]) +
                        ((itemPao.Count == 0) ? 0 : (itemPao.Count > 1) ?
                            Convert.ToInt32(itemPao[1].Ticket10) + Convert.ToInt32(itemPao[1].Ticket12) + Convert.ToInt32(itemPao[1].Ticket15) +
                            Convert.ToInt32(itemPao[0].Ticket10) + Convert.ToInt32(itemPao[0].Ticket12) + Convert.ToInt32(itemPao[0].Ticket15) :
                            Convert.ToInt32(itemPao[0].Ticket10) + Convert.ToInt32(itemPao[0].Ticket12) + Convert.ToInt32(itemPao[0].Ticket15))),
                    /*******************************************************************************************************************************/

                    SBikeTkt_num = Convert.ToString(Convert.ToInt32(SqlData[0, 9]) +
                        ((itemPao.Count == 0) ? 0 : (itemPao.Count > 1) ?
                            Convert.ToInt32(itemPao[1].TicketBike) + Convert.ToInt32(itemPao[0].TicketBike) :
                            Convert.ToInt32(itemPao[0].TicketBike))),
                    /*******************************************************************************************************************************/

                    OneDay1_num = Convert.ToString((Convert.ToInt32(SqlData[0, 10]) +
                        ((itemPao.Count == 0) ? 0 : (itemPao.Count > 1) ?
                            Convert.ToInt32(itemPao[1].Ticket_OneDay) + Convert.ToInt32(itemPao[1].ExchangeOneDay) +
                            Convert.ToInt32(itemPao[0].Ticket_OneDay) + Convert.ToInt32(itemPao[0].ExchangeOneDay) + Convert.ToInt32(itemPao[0].KKDay) + Convert.ToInt32(itemPao[0].Ntpc) + Convert.ToInt32(itemPao[0].SplitTicker) :
                            Convert.ToInt32(itemPao[0].Ticket_OneDay) + Convert.ToInt32(itemPao[0].ExchangeOneDay) + Convert.ToInt32(itemPao[0].KKDay) + Convert.ToInt32(itemPao[0].Ntpc) + Convert.ToInt32(itemPao[0].SplitTicker))) * 6),
                    /*******************************************************************************************************************************/

                    OneDay_Exit_num = Convert.ToString(Convert.ToInt32(SqlData[0, 11])),

                    /*******************************************************************************************************************************/

                    Lanwair1_num = Convert.ToString((Convert.ToInt32(SqlData[0, 12]) +
                        ((itemPao.Count == 0) ? 0 : (itemPao.Count > 1) ?
                            Convert.ToInt32(itemPao[1].Lanwair1) + Convert.ToInt32(itemPao[0].Lanwair1) :
                            Convert.ToInt32(itemPao[0].Lanwair1))) * 6),
                    /*******************************************************************************************************************************/

                    Lanwair2_num = Convert.ToString((Convert.ToInt32(SqlData[0, 13]) +
                        ((itemPao.Count == 0) ? 0 : (itemPao.Count > 1) ?
                            Convert.ToInt32(itemPao[1].Lanwair2) + Convert.ToInt32(itemPao[0].Lanwair2) :
                            Convert.ToInt32(itemPao[0].Lanwair2))) * 6),
                    /*******************************************************************************************************************************/

                    Group_Tik_num = Convert.ToString(Convert.ToInt32(SqlData[0, 14]) + Day_GroupTicket.Count),
                    /*******************************************************************************************************************************/

                    Group_Peo_num = Convert.ToString(Convert.ToInt32(SqlData[0, 15]) + GroupSumPeople),
                    /*******************************************************************************************************************************/

                    Office_Staff_num = Convert.ToString(Convert.ToInt32(SqlData[0, 16]))
                };

                MyCommond.WriteLog(ThisReceive, $"將輸出資料轉換為陣列");
                Data_Afc = MyCommond.TResultToArray<DayVolume>(ThisReceive, ExportData);
                MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_Afc, "Data_Afc");
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }

            ValueCount += 1;
            ValueCount = 80;
        }

        private void Calculator_Day_Volume(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            MethodeIsRun(LrtTxnClientSwitch.Day_Amount);
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");
            bool haveK0 = MainServer.globalCommond.Stock_Lrt_Station.FindAll(x => Convert.ToInt16(x.CodeNumber) > 150 & x.OperationLine == "Ak").Count > 0;
            bool runK0 = false;
            System.Threading.Thread thread = new System.Threading.Thread(() =>
            {
                runK0 = true;
                string cmdstr = @"--K0進站人數
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
            WHEN PeriodCode = 1  AND ValidPeriod   = 31 /*------------------*/ THEN '小月票'
            WHEN PeriodCode = 16 AND OriginalTimes = 1  /*------------------*/ THEN '退小月票'
            ELSE '效期票'
            END
        WHEN TxnType  IN (29,30) THEN 
            CASE 
            WHEN DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND ((IssueCode IN (2) AND PersonalProfile IN (5)) OR (IssueCode IN (9) AND CardType IN (7) AND AreaCode IN (21,22,23,25,30,31,32,33,35,36,37,39,40,41,43,44,45,47,48,49,51,52,53,55,56,57,59,60,61,63,64,65,67,68,69,71,72,73,75,76,77,79,80,81,83,84,85,87,88,89,91,92,93,95,96,97,99,100,101,103,104,105,107,108,109,111,112,113))) 
                THEN '行政院通勤月票_學生卡' 
            ELSE '行政院通勤月票_一般卡'
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
    UNION SELECT '出站'     AS TName, T1.LocationId,      T1.EquipType,              T1.CardSN, T1.TxnType, T1.TxnDT, T1.IssueCode,      T1.CardType,      T1.PersonalProfile,      T1.IdentityType,      T1.IdentityExpiryDT, T1.TxnAmt,      T1.StartLoc,      T1.EndLoc,  NULL AS SaleStartLoc,  NULL AS SaleEndLoc,      T1.TxnWP,      T1.TxnPoint,      T1.TrnAmt,      T1.AreaCode,      T1.WelfareAmt,      T1.PeriodCode,      T1.OriginalTimes,      T1.ValidPeriod,      T2.BodyData,      T1.LastTrnFlag FROM (SELECT * FROM Txn_Exit JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate) AS T1 LEFT JOIN (SELECT * FROM ECCTxn_Interface WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate) AS T2 ON T1.CardSN = T2.CardSN AND T1.TxnDT = T2.TxnDT AND T1.TxnType = T2.MyTxnType AND T1.EquipId = T2.EquipId 
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
/*-----(以上為站存檔，以下為主篩選分析區域)--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/

SELECT COUNT(*) AS 'K0進站' FROM #TempAllTxn WHERE TName IN ('進站') AND EquipType IN (7)

/*-----(以上為主篩選分析區域，以下為刪除暫存表)--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------*/
DROP TABLE #TempAllTxn
DROP TABLE #mTempLocation";
                var Station = (from p in MainServer.globalCommond.Stock_Lrt_Station select Convert.ToByte(p.CodeNumber)).ToList();
                var OperationData_Start = OperationDate.ToString("yyyy-MM-dd 05:00:00");
                var OperationData_End = OperationDate.AddDays(1).ToString("yyyy-MM-dd 05:00:00");
                var k0_data = ArrayToExport(SQLTextSelect(string.Format(cmdstr, Station.Min(), Station.Max(), OperationData_Start, OperationData_End)));

                string se = $@"{MainServer.OperationStartDate:yyyyMMdd},{k0_data[0, 0]}";

                string pa = $@"{MyCommond.Path_Program}{MyCommond.Path_Report}";
                int dy = MainServer.OperationStartDate.Year - 1911;
                string fa = MyCommond.CheckAndReName(pa, $@"{dy}{MainServer.OperationStartDate:MMdd}_K0進站_{MainServer.globalCommond.OperationCode_String}", "csv"); // 1130101_K0進站_安坑.csv
                MyCommond.ExportData(pa + fa, Encoding.UTF8, se);
            });
            thread.IsBackground = true;

            try
            {
                MyCommond.WriteLog(ThisReceive, $"開始電票相關資料擷取");
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) 
                {
                    MyCommond.WriteLog(ThisReceive, $"a"); 
                    Data_Afc = ArrayToExport(Calculator_1()); 
                    if (haveK0)
                    {
                        thread.Start();
                    }
                }
                else
                {
                    Data_Afc = ArrayToExport(inputZ(new string[1, new DayVolume().ExportTitle_zhTW().Split(',').Length]));
                }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_Afc, "Data_Afc");
                ValueCount = 40;

                MyCommond.WriteLog(ThisReceive, $"開始電支相關資料擷取");
                Data_MobilePay = new int[1, 1];
                Data_MobilePay[0, 0] = Day_MobilePayData.Count;
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_MobilePay, "Data_MobilePay");
                ValueCount = 60;

                MyCommond.WriteLog(ThisReceive, $"開始自有票相關資料擷取");
                Data_OwnTicket = new int[1,3];
                var item1 = MainServer.paoTicketSaleLists.FindAll(x => x.Operation == OperationDate.ToString("yyyyMMdd"));
                if (item1.Count == 0) { MyCommond.WriteLog(ThisReceive, "PaoData 異常"); }
                foreach (var itemPao in item1)
                {
                    Data_OwnTicket[0, 0] += (
                /*普通*/ (Convert.ToInt32(itemPao.Ticket20) + Convert.ToInt32(itemPao.Ticket25) + Convert.ToInt32(itemPao.Ticket30)) + 
                /*優待*/ (Convert.ToInt32(itemPao.Ticket10) + Convert.ToInt32(itemPao.Ticket12) + Convert.ToInt32(itemPao.Ticket15)) + 
                /*自行*/ (Convert.ToInt32(itemPao.TicketBike)) +
                /*兌換*/ (Convert.ToInt32(itemPao.ExchangeTicket20) + Convert.ToInt32(itemPao.ExchangeTicket25) + Convert.ToInt32(itemPao.ExchangeTicket30))
                        );
                    Data_OwnTicket[0, 1] += (
                          Convert.ToInt32(itemPao.Ticket_OneDay)
                        + Convert.ToInt32(itemPao.ExchangeOneDay)
                        + Convert.ToInt32((itemPao.SplitTicker != "") ? itemPao.SplitTicker : "0")
                        + Convert.ToInt32((itemPao.KKDay != "") ? itemPao.KKDay : "0")
                        + Convert.ToInt32((itemPao.Ntpc != "") ? itemPao.Ntpc : "0")
                        ) * 6;
                }
                Data_OwnTicket[0, 2] += Convert.ToInt32(Day_GroupTicket.Sum(x => Convert.ToInt16(x.People)).ToString());
                ValueCount = 80;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {
                if (runK0) thread.Join();
            }

        }

        #endregion

        #region 營運資料統計_營收

        private void Calculator_Day_Amount_Old(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            MethodeIsRun(LrtTxnClientSwitch.Day_Volume);
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                var MobilePay_CC = (from p in MainServer.globalCommond.Stock_MobileCardType where p.CardType == "Credit" select p.CardSubType).ToList();
                var MobilePay_QQ = (from p in MainServer.globalCommond.Stock_MobileCardType where p.CardType == "QrCode" select p.CardSubType).ToList();
                var Station_CodeNumber = (from p in MainServer.globalCommond.Stock_Lrt_Station select Convert.ToByte(p.CodeNumber)).ToList();

                var SqlFileName = MainServer.globalCommond.Stock_SqlList.Find(x => x.UseNameCh == LrtTxnClientSwitch.Day_Amount.ToDescription()).FileName1;
                var Sqlstr = File.ReadAllText(MyCommond.Path_Program + MyCreat.Path_Sql + SqlFileName + ".txt");
                var SqlCommand = string.Format(Sqlstr, Station_CodeNumber.Min(), Station_CodeNumber.Max(), OperationDate.ToString("yyyy/MM/dd 05:00:00"), OperationDate.AddDays(1).ToString("yyyy/MM/dd 05:00:00"));

                var SqlFileName_Od = MainServer.globalCommond.Stock_SqlList.Find(x => x.UseNameCh == LrtTxnClientSwitch.Day_OriginDestination.ToDescription()).FileName1;
                var Sqlstr_Od = File.ReadAllText(MyCommond.Path_Program + MyCreat.Path_Sql + SqlFileName_Od + ".txt");
                var SqlCommand_Od = string.Format(Sqlstr_Od, Station_CodeNumber.Min(), Station_CodeNumber.Max(), OperationDate.ToString("yyyy/MM/dd 05:00:00"), OperationDate.AddDays(1).ToString("yyyy/MM/dd 05:00:00"));

                var item1 = MainServer.paoTicketSaleLists.FindAll(x => x.Operation == OperationDate.ToString("yyyyMMdd")).ToList();
                if (item1.Count == 0) { MyCommond.WriteLog(ThisReceive, "PaoData 異常"); }
                ValueCount += 1;

                #region 電支
                MyCommond.WriteLog(ThisReceive, $"取得電子支付交易資料");
                var cc = (from p in Day_MobilePayData where MobilePay_CC.Contains(p.CardSubType) select p).ToList();
                var qq = (from p in Day_MobilePayData where MobilePay_QQ.Contains(p.CardSubType) select p).ToList();
                ValueCount += 1;

                MyCommond.WriteLog(ThisReceive, $"計算電子支付交易資料");
                int cc_a = 0, qq_a = 0;
                for (int i = 0; i < cc.Count; i++) { cc_a += Convert.ToInt32(cc[i].RealPay); }
                for (int i = 0; i < qq.Count; i++) { qq_a += Convert.ToInt32(qq[i].RealPay); }
                ValueCount += 1;

                MyCommond.WriteLog(ThisReceive, $"計算電子支付OD");
                int SubsidyAmt = 0;
                foreach (var item in Day_MobilePayData)
                {
                    string Station_I = MainServer.globalCommond.Stock_Lrt_Station.Find(x => x.StationName == item.EntryStation).CodeName;
                    string Station_O = MainServer.globalCommond.Stock_Lrt_Station.Find(x => x.StationName == item.ExitStation).CodeName;
                    var sss = MainServer.globalCommond.Stock_Lrt_Subsidy.Find(x => x.Operation_Area == MainServer.globalCommond.OperationCode && x.StartStation == Station_I && x.EndStation == Station_O).OdAmt;
                    SubsidyAmt += Convert.ToInt32(sss);
                    MyCommond.WriteLog(ThisReceive, $"進站:{Station_I,4}, 出站:{Station_O,4}, 補貼:{sss,4}, 比數:{1,8}, 累計補貼金額:{SubsidyAmt}");
                }
                ValueCount += 1;

                #endregion
                ValueCount = 40;

                #region 電票
                MyCommond.WriteLog(ThisReceive, $"取得電票OD");
                string[,] SqlOd = inputZ(new string[Station_CodeNumber.Count, Station_CodeNumber.Count + 1]);
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); SqlOd = SQLTextSelect(SqlCommand_Od); }

                MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), SqlOd, "SqlOd");

                ValueCount += 1;

                MyCommond.WriteLog(ThisReceive, $"計算電票OD");

                for (int i = 0; i < SqlOd.GetLength(0); i++)
                {
                    for (int j = 1; j < SqlOd.GetLength(1); j++)
                    {
                        string Station_I = "";
                        string Station_O = "";
                        string TicketSubsidy = "";
                        int DataCount = 0;
                        try
                        {
                            Station_I = MainServer.globalCommond.Stock_Lrt_Station[i].CodeName;
                            Station_O = MainServer.globalCommond.Stock_Lrt_Station[j - 1].CodeName;
                            TicketSubsidy = MainServer.globalCommond.Stock_Lrt_Subsidy.Find(x => x.Operation_Area == MainServer.globalCommond.OperationCode && x.StartStation == Station_I && x.EndStation == Station_O).OdAmt;
                            DataCount = (SqlOd[i, j] != null || SqlOd[i, j] == "") ? Convert.ToInt32(SqlOd[i, j]) : 0;

                            //MyCommond.WriteLog(ThisReceive, $"進站:{Station_I,4}, 出站:{Station_O,4}, 補貼:{TicketSubsidy,4}, 比數:{DataCount,8}");

                            if (TicketSubsidy != null && Convert.ToInt32(TicketSubsidy) > 0) { SubsidyAmt += (Convert.ToInt32(TicketSubsidy) * DataCount); }
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error\n{ex.Message}");
                        }
                        finally
                        {
                            MyCommond.WriteLog(ThisReceive, $"進站:{Station_I,4}, 出站:{Station_O,4}, 補貼:{TicketSubsidy,4}, 比數:{DataCount,8}, 累計補貼金額:{SubsidyAmt}");
                        }
                    }
                }

                ValueCount += 1;

                MyCommond.WriteLog(ThisReceive, $"取得電票交易資料");
                string[] mTitle = new DayAmount().ExportTitle_zhTW().Split(',');
                string[,] SqlData = inputZ(new string[1, new DayAmount().ExportTitle_zhTW().Split(',').Length]);
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); SqlData = SQLTextSelect(SqlCommand); }

                MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), SqlData, "SqlData");

                ValueCount += 1;

                #endregion
                ValueCount = 60;
                MyCommond.WriteLog(ThisReceive, $"撰寫輸出資料");
                DayAmount ExportData = new DayAmount()
                {
                    Electront_Amt = Convert.ToString(Convert.ToInt32(SqlData[0, 0])),
                    Student_Amt = Convert.ToString(Convert.ToInt32(SqlData[0, 1])),
                    Welfare_Amt = Convert.ToString(Convert.ToInt32(SqlData[0, 2])),
                    Credit_Amt = Convert.ToString(cc_a),
                    Qrcode_Amt = Convert.ToString(qq_a),
                    /*******************************************************************************************************************************/

                    SOneTkt_Amt = Convert.ToString(Convert.ToInt32(SqlData[0, 5]) + ((item1.Count == 0) ? 0 : (item1.Count > 1) ?
                                (Convert.ToInt32(item1[1].Ticket20) * 20 + Convert.ToInt32(item1[1].Ticket25) * 25 + Convert.ToInt32(item1[1].Ticket30) * 30 +
                                 Convert.ToInt32(item1[0].Ticket20) * 20 + Convert.ToInt32(item1[0].Ticket25) * 25 + Convert.ToInt32(item1[0].Ticket30) * 30) :
                                (Convert.ToInt32(item1[0].Ticket20) * 20 + Convert.ToInt32(item1[0].Ticket25) * 25 + Convert.ToInt32(item1[0].Ticket30) * 30))),
                    /*******************************************************************************************************************************/

                    SConcessionTkt_Amt = Convert.ToString(Convert.ToInt32(SqlData[0, 6]) + ((item1.Count == 0) ? 0 : (item1.Count > 1) ?
                                (Convert.ToInt32(item1[1].Ticket10) * 10 + Convert.ToInt32(item1[1].Ticket12) * 12 + Convert.ToInt32(item1[1].Ticket15) * 15 - Convert.ToInt32(item1[1].ConcessionRefund) +
                                 Convert.ToInt32(item1[0].Ticket10) * 10 + Convert.ToInt32(item1[0].Ticket12) * 12 + Convert.ToInt32(item1[0].Ticket15) * 15) - Convert.ToInt32(item1[0].ConcessionRefund) :
                                (Convert.ToInt32(item1[0].Ticket10) * 10 + Convert.ToInt32(item1[0].Ticket12) * 12 + Convert.ToInt32(item1[0].Ticket15) * 15) - Convert.ToInt32(item1[0].ConcessionRefund))),
                    /*******************************************************************************************************************************/

                    SBikeTkt_Amt = Convert.ToString(Convert.ToInt32(SqlData[0, 7]) + ((item1.Count == 0) ? 0 : (item1.Count > 1) ?
                                (Convert.ToInt32(item1[1].TicketBike) * 50 +
                                 Convert.ToInt32(item1[0].TicketBike) * 50) :
                                (Convert.ToInt32(item1[0].TicketBike) * 50))),
                    /*******************************************************************************************************************************/

                    OneDay1_Amt = Convert.ToString(Convert.ToInt32(SqlData[0, 8]) + ((item1.Count == 0) ? 0 : (item1.Count > 1) ?
                                (Convert.ToInt32(item1[1].Ticket_OneDay) * 50 +
                                 Convert.ToInt32(item1[0].Ticket_OneDay) * 50) :
                                (Convert.ToInt32(item1[0].Ticket_OneDay) * 50))),
                    /*******************************************************************************************************************************/

                    OneDay_Exit_Amt = Convert.ToString(Convert.ToInt32(SqlData[0, 9])),

                    /*******************************************************************************************************************************/

                    Lanwair1_Amt = Convert.ToString(Convert.ToInt32(SqlData[0, 10]) + ((item1.Count == 0) ? 0 : (item1.Count > 1) ?
                                (Convert.ToInt32(item1[1].Lanwair1) * 50 +
                                 Convert.ToInt32(item1[0].Lanwair1) * 50) :
                                (Convert.ToInt32(item1[0].Lanwair1) * 50))),
                    /*******************************************************************************************************************************/

                    Lanwair2_Amt = Convert.ToString(Convert.ToInt32(SqlData[0, 11]) + ((item1.Count == 0) ? 0 : (item1.Count > 1) ?
                                (Convert.ToInt32(item1[1].Lanwair2) * 50 +
                                 Convert.ToInt32(item1[0].Lanwair2) * 50) :
                                (Convert.ToInt32(item1[0].Lanwair2) * 50))),
                    /*******************************************************************************************************************************/

                    Group_Amt = Convert.ToString(Convert.ToInt32(SqlData[0, 12]) + ((item1.Count == 0) ? 0 : (item1.Count > 1) ?
                                (Convert.ToInt32(item1[1].GroupAmt) +
                                 Convert.ToInt32(item1[0].GroupAmt)) :
                                (Convert.ToInt32(item1[0].GroupAmt)))),
                    /*******************************************************************************************************************************/

                    Lanwair4_Amt = Convert.ToString(Convert.ToInt32(SqlData[0, 13]) + ((item1.Count == 0) ? 0 : (item1.Count > 1) ?
                                (Convert.ToInt32(item1[1].OpenTrafficSetTicket) * 450 +
                                 Convert.ToInt32(item1[0].OpenTrafficSetTicket) * 450) :
                                (Convert.ToInt32(item1[0].OpenTrafficSetTicket) * 450))),
                    /*******************************************************************************************************************************/

                    Retik_Amt = Convert.ToString(Convert.ToInt32(SqlData[0, 14])),
                    ConcessionTicket = Convert.ToString(Convert.ToInt32(SqlData[0, 15])),
                    ConcessionPoint = Convert.ToString(Convert.ToInt32(SqlData[0, 16])),
                    SOneConcessionReBack = Convert.ToString(Convert.ToInt32(SqlData[0, 17])),
                    YouBickToLrt = Convert.ToString(Convert.ToInt32(SqlData[0, 18])),
                    AllPassPayback = Convert.ToString(Convert.ToInt32(SqlData[0, 19])),
                    RideSubsidy = Convert.ToString(SubsidyAmt),
                };
                ValueCount += 1;

                MyCommond.WriteLog(ThisReceive, $"將輸出資料轉換為陣列");
                Data_Afc = MyCommond.TResultToArray<DayAmount>(ThisReceive, ExportData);
                ValueCount += 1;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }

            ValueCount = 80;
        }

        private void Calculator_Day_Amount(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            MethodeIsRun(LrtTxnClientSwitch.Day_Volume);
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                int SubsidyAmt = 0;

                MyCommond.WriteLog(ThisReceive, $"開始電票相關資料擷取");
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); Data_Afc = ArrayToExport(Calculator_1()); }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_Afc, "Data_Afc");

                MyCommond.WriteLog(ThisReceive, $"取得電票OD");
                var Station_CodeNumber = (from p in MainServer.globalCommond.Stock_Lrt_Station select Convert.ToByte(p.CodeNumber)).ToList();
                var SqlFileName_Od = MainServer.globalCommond.Stock_SqlList.Find(x => x.UseNameCh == LrtTxnClientSwitch.Day_OriginDestination.ToDescription()).FileName1;
                var Sqlstr_Od = File.ReadAllText(MyCommond.Path_Program + MyCreat.Path_Sql + SqlFileName_Od + ".txt");
                var SqlCommand_Od = string.Format(Sqlstr_Od, Station_CodeNumber.Min(), Station_CodeNumber.Max(), OperationDate.ToString("yyyy/MM/dd 05:00:00"), OperationDate.AddDays(1).ToString("yyyy/MM/dd 05:00:00"));
                string[,] SqlOd = inputZ(new string[Station_CodeNumber.Count, Station_CodeNumber.Count + 1]);
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); SqlOd = SQLTextSelect(SqlCommand_Od); }

                MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), SqlOd, "SqlOd");

                ValueCount += 1;

                MyCommond.WriteLog(ThisReceive, $"計算電票OD");

                for (int i = 0; i < SqlOd.GetLength(0); i++)
                {
                    for (int j = 1; j < SqlOd.GetLength(1); j++)
                    {
                        string Station_I = "";
                        string Station_O = "";
                        string TicketSubsidy = "";
                        int DataCount = 0;
                        try
                        {
                            Station_I = MainServer.globalCommond.Stock_Lrt_Station[i].CodeName;
                            Station_O = MainServer.globalCommond.Stock_Lrt_Station[j - 1].CodeName;
                            TicketSubsidy = MainServer.globalCommond.Stock_Lrt_Subsidy.Find(x => x.Operation_Area == MainServer.globalCommond.OperationCode && x.StartStation == Station_I && x.EndStation == Station_O).OdAmt;
                            DataCount = (SqlOd[i, j] != null || SqlOd[i, j] == "") ? Convert.ToInt32(SqlOd[i, j]) : 0;

                            //MyCommond.WriteLog(ThisReceive, $"進站:{Station_I,4}, 出站:{Station_O,4}, 補貼:{TicketSubsidy,4}, 比數:{DataCount,8}");

                            if (TicketSubsidy != null && Convert.ToInt32(TicketSubsidy) > 0) { SubsidyAmt += (Convert.ToInt32(TicketSubsidy) * DataCount); }
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error\n{ex.Message}");
                        }
                        finally
                        {
                            MyCommond.WriteLog(ThisReceive, $"進站:{Station_I,4}, 出站:{Station_O,4}, 補貼:{TicketSubsidy,4}, 比數:{DataCount,8}, 累計補貼金額:{SubsidyAmt}");
                        }
                    }
                }

                ValueCount = 40;

                MyCommond.WriteLog(ThisReceive, $"開始電支相關資料擷取");
                Data_MobilePay = new int[1, 1];
                Data_MobilePay[0,0] = Day_MobilePayData.Sum(x => x.RealPay);
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_MobilePay, "Data_MobilePay");
                MyCommond.WriteLog(ThisReceive, $"計算電子支付OD");
                foreach (var item in Day_MobilePayData)
                {
                    string Station_I = MainServer.globalCommond.Stock_Lrt_Station.Find(x => x.StationName == item.EntryStation).CodeName;
                    string Station_O = MainServer.globalCommond.Stock_Lrt_Station.Find(x => x.StationName == item.ExitStation).CodeName;
                    var sss = MainServer.globalCommond.Stock_Lrt_Subsidy.Find(x => x.Operation_Area == MainServer.globalCommond.OperationCode && x.StartStation == Station_I && x.EndStation == Station_O).OdAmt;
                    SubsidyAmt += Convert.ToInt32(sss);
                    MyCommond.WriteLog(ThisReceive, $"進站:{Station_I,4}, 出站:{Station_O,4}, 補貼:{sss,4}, 比數:{1,8}, 累計補貼金額:{SubsidyAmt}");
                }
                ValueCount = 60;

                MyCommond.WriteLog(ThisReceive, $"開始自有票相關資料擷取");
                Data_OwnTicket = new int[1, 4];
                var item1 = MainServer.paoTicketSaleLists.FindAll(x => x.Operation == OperationDate.ToString("yyyyMMdd"));
                if (item1.Count == 0) { MyCommond.WriteLog(ThisReceive, "PaoData 異常"); }

                if (true)
                {
                    MyCommond.WriteLog(ThisReceive, $@"符合條件的PAO資料數量:{item1.Count}");
                    MyCommond.WriteLog(ThisReceive, $@"{new PaoTicketSaleList().ExportTitle_zhTW()}");
                    MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), item1, "");
                }

                try
                {
                    foreach (var itemPao in item1)
                    {

                        var test1 = (Convert.ToInt32(itemPao.Ticket20) * 20 + Convert.ToInt32(itemPao.Ticket25) * 25 + Convert.ToInt32(itemPao.Ticket30) * 30);
                        var test2 = (Convert.ToInt32(itemPao.Ticket10) * 10 + Convert.ToInt32(itemPao.Ticket12) * 12 + Convert.ToInt32(itemPao.Ticket15) * 15 - Convert.ToInt32(itemPao.ConcessionRefund));
                        var test3 = (Convert.ToInt32(itemPao.TicketBike));
                        var test4 = 0;


                        Data_OwnTicket[0, 0] += (
                    /*普通*/ (Convert.ToInt32(itemPao.Ticket20) * 20 + Convert.ToInt32(itemPao.Ticket25) * 25 + Convert.ToInt32(itemPao.Ticket30) * 30)
                    /*優待*/ + (Convert.ToInt32(itemPao.Ticket10) * 10 + Convert.ToInt32(itemPao.Ticket12) * 12 + Convert.ToInt32(itemPao.Ticket15) * 15 - Convert.ToInt32(itemPao.ConcessionRefund)) // 系統售票(未退) + 人工售票(已退) - 人工退半價
                    /*自行*/ + (Convert.ToInt32(itemPao.TicketBike) * 50)
                            );
                        Data_OwnTicket[0, 1] += (
                            (Convert.ToInt32(itemPao.Ticket_OneDay) * 50)
                            + (Convert.ToInt32(itemPao.Lanwair1) * 50)
                            + (Convert.ToInt32(itemPao.Lanwair2) * 50)
                            );
                        Data_OwnTicket[0, 2] += (
                            Convert.ToInt32(itemPao.GroupAmt ?? "0")
                            + (Convert.ToInt32(itemPao.OpenTrafficSetTicket) * 450)
                            );
                        Data_OwnTicket[0, 3] += (
                            Convert.ToInt32(itemPao.ConcessionRefund)
                            + (Convert.ToInt32(itemPao.Ticket10) * 10 + Convert.ToInt32(itemPao.Ticket12) * 13 + Convert.ToInt32(itemPao.Ticket15) * 15)
                            );
                    }


                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"ERROR!訊息:{ex.Message}");
                }
                // MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_OwnTicket, "Data_OwnTicket");

                MyCommond.WriteLog(ThisReceive, $"寫入OD");
                Data_Afc[Data_Afc.GetLength(0) - 1, Data_Afc.GetLength(1) - 1] += SubsidyAmt;

                ValueCount = 80;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }

        }

        #endregion

        #region 每日全線各站分時運量日報表

        private void Calculator_Day_EachStation_EachTime(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                //List<LrtStationList> tempStation = MainServer.globalCommond.Stock_Lrt_Station.GroupBy(g => g.CodeNumber);
                //List<LrtStationList> tempStation = (from p in MainServer.globalCommond.Stock_Lrt_Station group p by new { p.OperationLine, p.CodeNumber, p.CodeName } into g select new LrtStationList {g.Key.OperationLine, g.Key.CodeNumber, g.Key.CodeName}) ;
                List<string> tempStation = (from p in MainServer.globalCommond.Stock_Lrt_Station group p by new { p.CodeNumber } into g select g.Key.CodeNumber).ToList();

                Data_Afc = new int[24 + 1, tempStation.Count * 2];
                Data_MobilePay = new int[24 + 1, tempStation.Count * 2];
                Data_OwnTicket = new int[24 + 1, tempStation.Count * 2];
                ValueCount += 1;

                int temp_Day_Hour = 24;
                int temp_Start_Hour = 5;
                #region 人工
                var item1 = MainServer.paoTicketSaleLists.FindAll(x => x.Operation == OperationDate.ToString("yyyyMMdd"));
                if (item1.Count == 0) { MyCommond.WriteLog(ThisReceive, "PaoData 異常"); }
                MyCommond.WriteLog(ThisReceive, $"將人工票寫入分時分站");
                foreach (var item in item1)
                {
                    var Offset_Row = Data_OwnTicket.GetLength(0) - 1;

                    var aa = MainServer.globalCommond.Stock_Lrt_Station.Find(x => x.CodeName == item.Station);

                    var Offset_Column = tempStation.FindIndex(x => x == aa.CodeNumber);

                    var t10 = Convert.ToInt16(item.Ticket10);
                    var t12 = Convert.ToInt16(item.Ticket12);
                    var t15 = Convert.ToInt16(item.Ticket15);
                    var t20 = Convert.ToInt16(item.Ticket20);
                    var t25 = Convert.ToInt16(item.Ticket25);
                    var t30 = Convert.ToInt16(item.Ticket30);
                    var tb = Convert.ToInt16(item.TicketBike);
                    var et20 = Convert.ToInt16(item.ExchangeTicket20);
                    var et25 = Convert.ToInt16(item.ExchangeTicket25);
                    var et30 = Convert.ToInt16(item.ExchangeTicket30);
                    var tod = Convert.ToInt16(item.Ticket_OneDay);
                    var eod = Convert.ToInt16(item.ExchangeOneDay);
                    var st = Convert.ToInt16((item.SplitTicker != "") ? item.SplitTicker : "0");
                    var kd = Convert.ToInt16((item.KKDay != "") ? item.KKDay : "0");
                    var nt = Convert.ToInt16((item.Ntpc != "") ? item.Ntpc : "0");
                    var l1 = Convert.ToInt16(item.Lanwair1);
                    var l2 = Convert.ToInt16(item.Lanwair2);

                    var Entry_Ride = t10 + t12 + t15 + t20 + t25 + t30 + tb + et20 + et25 + et30 + (tod + eod + st + kd + nt + l1 + l2) * 6;
                    var Exit_Ride = Entry_Ride;

                    string str = $"\n" +
                        $"人工車站:{item.Station}\n" +
                        $"單程票10元:{t10,4}張\n" +
                        $"單程票12元:{t12,4}張\n" +
                        $"單程票15元:{t15,4}張\n" +
                        $"單程票20元:{t20,4}張\n" +
                        $"單程票25元:{t25,4}張\n" +
                        $"單程票30元:{t30,4}張\n" +
                        $"自行車票  :{tb,4}張\n" +
                        $"兌換  20元:{et20,4}張\n" +
                        $"兌換  25元:{et25,4}張\n" +
                        $"兌換  30元:{et30,4}張\n" +
                        $"    一日票:{tod,4}張\n" +
                        $"兌換一日票:{eod,4}張\n" +
                        $"   切   票:{st,4}張\n" +
                        $"     KKDAY:{kd,4}張\n" +
                        $"      NTPC:{nt,4}張\n" +
                        $"    活動1 :{l1,4}張\n" +
                        $"    活動2 :{l2,4}張\n" +
                        $"共計{Entry_Ride,8}人次\n";
                    MyCommond.WriteLog(ThisReceive, str);

                    Data_OwnTicket[Offset_Row, Offset_Column * 2 + 0] += Entry_Ride;
                    Data_OwnTicket[Offset_Row, Offset_Column * 2 + 1] += Exit_Ride;
                }
                ValueCount += 1;

                MyCommond.WriteLog(ThisReceive, $"將團體票寫入分時分站");
                string GroupPathString = $"{MainServer.Path_PaoData}{OperationDate.ToString("yyyyMMdd")}_{MainServer.FileName_GroupTicket}_{MainServer.globalCommond.OperationCode_String}.csv";
                var item_group = MyCommond.ToViewModel<GroupTicketList>(ThisReceive, GroupPathString, new GroupTicketList().ExportTitle_zhTW(), 1, true);
                foreach (var item in item_group)
                {
                    var Offset_Row = Convert.ToInt16(item.SaleTime.Split(':')[0]) - 5;
                    if (item.SaleDate != OperationDate.ToString("yyyyMMdd")) { Offset_Row = 5; }

                    var Offset_Column_Entry = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.CodeName == item.Entry);
                    var Offset_Column_Exit_ = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.CodeName == item.Exit);
                    var ItemRide = item.People;

                    string str = $"\n" +
                        $"進站座標行:{Offset_Row,4},欄:{Offset_Column_Entry,4}寫入{ItemRide}\n" +
                        $"出站座標行:{Offset_Row,4},欄:{Offset_Column_Exit_,4}寫入{ItemRide}\n";
                    MyCommond.WriteLog(ThisReceive, str);

                    Data_OwnTicket[Offset_Row, Offset_Column_Entry * 2 + 0] += Convert.ToInt16(ItemRide);
                    Data_OwnTicket[Offset_Row, Offset_Column_Exit_ * 2 + 1] += Convert.ToInt16(ItemRide);
                }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_OwnTicket, "Data_OwnTicket");
                #endregion
                ValueCount = 40;
                #region 電支
                var mp = Day_MobilePayData.ToList();
                MyCommond.WriteLog(ThisReceive, $"將電支寫入分時分站");
                foreach (var item in mp)
                {
                    var temp_Hour_Entry = Convert.ToInt16(item.EntryTime.ToString("HH"));
                    var Offset_Row_Entry = (temp_Hour_Entry - temp_Start_Hour >= 0) ? temp_Hour_Entry - temp_Start_Hour : temp_Day_Hour + temp_Hour_Entry - temp_Start_Hour;
                    var Offset_Column_Entry = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == item.EntryStation);

                    var temp_Hour_Exit_ = Convert.ToInt16(item.ExitTime.ToString("HH"));
                    var Offset_Row_Exit_ = (temp_Hour_Exit_ - temp_Start_Hour >= 0) ? temp_Hour_Exit_ - temp_Start_Hour : temp_Day_Hour + temp_Hour_Exit_ - temp_Start_Hour;
                    var Offset_Column_Exit_ = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == item.ExitStation);

                    Data_MobilePay[Offset_Row_Entry, Offset_Column_Entry * 2 + 0] += 1;
                    Data_MobilePay[Offset_Row_Exit_, Offset_Column_Exit_ * 2 + 1] += 1;
                }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_MobilePay, "Data_MobilePay");
                #endregion
                ValueCount = 60;
                #region 電票
                MyCommond.WriteLog(ThisReceive, $"取得電票資料");
                string[,] Data = null;
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); Data = Calculator_1(); }
                if (Data == null) return;
                MyCommond.WriteLog(ThisReceive, $"資料行數:{Data.GetLength(0)}");
                MyCommond.WriteLog(ThisReceive, $"將電票寫入分時分站");
                for (int i = 0; i < Data.GetLength(0); i++)
                {
                    // 0-Year, 1-Month, 2-Day, 3-Hour, 4-mCheck, 5-mCheckOffset, 6-mCheckIO, 7-LocationId, 8-num 
                    // 2023,   9,       19,    11,     Entry,    0,              +,          101,          200,
                    string ttt = "";
                    for (int j = 0; j < Data.GetLength(1); j++)
                    {
                        ttt += Data[i, j].ToString() + ",";
                    }
                    MyCommond.WriteLog(ThisReceive, $"Data[{i,4}] = {ttt}");

                    var temp_Hour = Convert.ToInt16(Data[i, 3]);
                    MyCommond.WriteLog(ThisReceive, $"Data[{i}, 3] = {Data[i, 3]}");
                    var Offset_Row = (temp_Hour - temp_Start_Hour >= 0) ? temp_Hour - temp_Start_Hour : temp_Day_Hour + temp_Hour - temp_Start_Hour;
                    MyCommond.WriteLog(ThisReceive, $"Offset_Row    = {Offset_Row}");
                    var Offset_Column = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.CodeNumber == Convert.ToString(Data[i, 7]));
                    MyCommond.WriteLog(ThisReceive, $"Offset_Column = {Offset_Column}");
                    var Offset_IO = Convert.ToInt16(Data[i, 5]);
                    MyCommond.WriteLog(ThisReceive, $"Offset_IO = {Offset_IO}");
                    var Count_Num = Convert.ToInt16(Data[i, 8]);
                    MyCommond.WriteLog(ThisReceive, $"Count_Num = {Count_Num}");
                    var temp = Data_Afc[Offset_Row, Offset_Column * 2 + Offset_IO];
                    MyCommond.WriteLog(ThisReceive, $"temp = {temp}");
                    var tempString = (Data[i, 6] == "+") ? Convert.ToString(temp + Count_Num) : ((Data[i, 6] == "-") ? Convert.ToString(temp - Count_Num) : "ERROR");
                    MyCommond.WriteLog(ThisReceive, $"ExportData[{Offset_Row,4}, {Offset_Column,4} * 2 + {Offset_IO,4}] = {temp,8} {Data[i, 6]} {Count_Num,8} = {tempString}");
                    /**/
                    if (Data[i, 6] == "+") Data_Afc[Offset_Row, Offset_Column * 2 + Offset_IO] += Count_Num;
                    else if (Data[i, 6] == "-") Data_Afc[Offset_Row, Offset_Column * 2 + Offset_IO] -= Count_Num;
                    MyCommond.WriteLog(ThisReceive, $"**********************************************");
                }
                #endregion
                ValueCount = 80;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }
        }

        #endregion

        #region 票卡進出站運量日報表

        private void Calculator_Day_ElectronicTicket(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                MyCommond.WriteLog(ThisReceive, $"開始相關電票資料擷取");
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); Data_Afc = ArrayToExport(Calculator_1()); }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_Afc, "Data_Afc");
                ValueCount = 40;

                MyCommond.WriteLog(ThisReceive, $"開始相關電支資料擷取");
                Data_MobilePay = Volume_Mobile_2_2_Line(Day_MobilePayData, ReportSwitch_1.Station);
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_MobilePay, "Data_MobilePay");
                ValueCount = 60;

                MyCommond.WriteLog(ThisReceive, $"無須填入團體票");
                Data_OwnTicket = null;
                ValueCount += 1;
                ValueCount = 80;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }
        }

        #endregion

        #region 日運量統計表(停用)

        private void Calculator_Day_AllRideList(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); Data_Afc = ArrayToExport(Calculator_1()); }

                Data_MobilePay = null;


                Data_OwnTicket = null;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }
        }

        #endregion

        #region 營收日報表

        private void Calculator_Day_TrafficAmount(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                MyCommond.WriteLog(ThisReceive, $"開始相關電票資料擷取");
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); Data_Afc = ArrayToExport(Calculator_1()); }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_Afc, "Data_Afc");
                ValueCount = 40;

                MyCommond.WriteLog(ThisReceive, $"開始相關電支資料擷取");
                Data_MobilePay = Amount_Mobile_2_Line(Day_MobilePayData, ReportSwitch_1.Station);
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_MobilePay, "Data_MobilePay");
                ValueCount = 60;

                MyCommond.WriteLog(ThisReceive, $"開始相關團體票擷取");
                var fsf = (from p in MainServer.paoTicketSaleLists where p.Operation == OperationDate.ToString("yyyyMMdd") select p).ToList();
                if (fsf.Count == 0) { MyCommond.WriteLog(ThisReceive, "PaoData 異常"); }

                int aaa = MainServer.globalCommond.Stock_Lrt_Station.Count;
                Data_OwnTicket = new int[aaa, 12];
                foreach(var item in fsf)
                {
                    int row = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.CodeName == item.Station);
                    MyCommond.WriteLog(ThisReceive, $"row -> {row,4}, {item.ToString()}");
                    Data_OwnTicket[row,  0] = Convert.ToInt32(item.Ticket20) * 20 + Convert.ToInt32(item.Ticket25) * 25 + Convert.ToInt32(item.Ticket30) * 30;
                    Data_OwnTicket[row,  1] = Convert.ToInt32(item.Ticket10) * 10 + Convert.ToInt32(item.Ticket12) * 12 + Convert.ToInt32(item.Ticket15) * 15;
                    Data_OwnTicket[row,  2] = Convert.ToInt32(item.TicketBike) * 50;
                    Data_OwnTicket[row,  3] = Convert.ToInt32(item.Ticket_OneDay) * 50;
                    Data_OwnTicket[row,  4] = Convert.ToInt32(item.Lanwair1) * 50;
                    Data_OwnTicket[row,  5] = Convert.ToInt32(item.GroupAmt);
                    Data_OwnTicket[row,  6] = Convert.ToInt32(item.OpenTrafficSetTicket) * 450;
                    Data_OwnTicket[row,  7] = Convert.ToInt32(item.RepairTicket);
                    Data_OwnTicket[row,  8] = Convert.ToInt32(item.AfcTakeOut);
                    Data_OwnTicket[row,  9] = Convert.ToInt32(item.VavmRefund);
                    Data_OwnTicket[row, 10] = Convert.ToInt32(item.SaleCardRefund);
                    Data_OwnTicket[row, 11] = Convert.ToInt32(item.PvRefund);
                }

                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_OwnTicket, "Data_OwnTicket");
                ValueCount = 80;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }
        }

        #endregion

        #region 每日起訖總表

        private void Calculator_Day_OriginDestination(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                var station_ = MainServer.globalCommond.Stock_Lrt_Station.ToList();
                int stationCount = station_.Count;
                ValueCount += 1;

                MyCommond.WriteLog(ThisReceive, $"開始相關電票資料擷取");
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); Data_Afc = ArrayToExport(Calculator_1()); }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_Afc, "Data_Afc");
                ValueCount = 40;

                MyCommond.WriteLog(ThisReceive, $"開始相關電支資料擷取");
                Data_MobilePay = null;
                foreach (var item in Day_MobilePayData)
                {
                    if (item.Volume == MobilePayVolum.InTime) { continue; }
                    int EntryStation = station_.FindIndex(x => x.StationName == item.EntryStation);
                    int ExitStation_ = station_.FindIndex(x => x.StationName == item.ExitStation);
                    Data_Afc[EntryStation, ExitStation_]++;
                }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_MobilePay, "Data_MobilePay");
                ValueCount = 60;

                MyCommond.WriteLog(ThisReceive, $"開始相關團票資料擷取");
                Data_OwnTicket = null;
                foreach (var item in Day_GroupTicket)
                {
                    int EntryStation = station_.FindIndex(x => x.CodeName == item.Entry);
                    int ExitStation_ = station_.FindIndex(x => x.CodeName == item.Exit);
                    Data_Afc[EntryStation, ExitStation_] += Convert.ToInt32(item.People);
                }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_OwnTicket, "Data_OwnTicket");
                ValueCount = 80;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }
        }

        #endregion

        #region 設備營收日報表

        private void Calculator_Day_EquipAmount(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                MyCommond.WriteLog(ThisReceive, $"開始相關電票資料擷取");
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); Data_Afc = ArrayToExport(Calculator_1()); }
                ValueCount = 40;

                MyCommond.WriteLog(ThisReceive, $"開始相關電支資料擷取");
                Data_MobilePay = Amount_10_Line(Day_MobilePayData, ReportSwitch_1.Station);
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_MobilePay, "Data_MobilePay");
                ValueCount = 60;

                MyCommond.WriteLog(ThisReceive, $"無須填入團體票");
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_OwnTicket, "Data_OwnTicket");
                ValueCount = 80;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }
        }

        #endregion

        #endregion

        #region 月

        List<MobilePay> Month_MobilePayData;
        List<GroupTicketList> Month_GroupTickets;

        private bool DataLoad_Month()
        {
            bool tf = false;
            Month_MobilePayData = new List<MobilePay>();
            Month_GroupTickets = new List<GroupTicketList>();

            Int32 DayCount = (
                    Convert.ToDateTime(MainServer.OperationEndDate.ToString("yyyy/MM/dd")) -
                    Convert.ToDateTime(MainServer.OperationStartDate.ToString("yyyy/MM/dd"))
                ).Days + 1;

            ValueCount += 1;

            #region 電支

            string MobilePayPath = $"{MyCommond.Path_Program + MyCommond.Path_Analyze}MobileData_{MainServer.OperationStartDate.ToString("yyyy-MM")}.csv";
            MyCommond.WriteLog(ThisReceive, $"本次撈取電支路徑:{MobilePayPath}");
            Month_MobilePayData = MainServer.ReadMobilePayData.FindAll(x => 
                   x.OpeLine == MainServer.globalCommond.OperationCode_String + "輕軌" 
                && MainServer.OperationStartDate <= x.ExitTime 
                && x.ExitTime < MainServer.OperationEndDate
                && x.Volume != MobilePayVolum.InTime
                ).ToList();

            ListDataCount(Month_MobilePayData);

            #endregion

            ValueCount += 1;

            #region 團體票 (只能算運量)

            for (int i = 0; i < DayCount - 1; i++)
            {
                var Ride_Date = MainServer.OperationStartDate.AddDays(i).ToString("yyyyMMdd");
                string GroupTicketPath = $"{MyCommond.Path_Program}{MainServer.Path_PaoData}" +
                $"{Ride_Date}_{MainServer.FileName_GroupTicket}_{MainServer.globalCommond.OperationCode_String}.csv";
                MyCommond.WriteLog(ThisReceive, $"本次讀取團體票路徑:{GroupTicketPath}");
                if (MyCommond.CheckFile(GroupTicketPath))
                {
                LoadData:
                    List<GroupTicketList> item = MyCommond.ToViewModel<GroupTicketList>(ThisReceive, GroupTicketPath, new GroupTicketList().ExportTitle_zhTW(), 1);
                    bool errorData = false;

                    foreach (var jtem in item)
                    {
                        if (!errorData && jtem.SaleStation == null)
                        {
                            errorData = true;
                            break;
                        }
                        MyCommond.WriteLog(ThisReceive, $"讀取的資料:{jtem.ToString()}");
                        if (jtem.SaleDate != Ride_Date)
                        {
                            MyCommond.WriteLog(ThisReceive, $"非購買日使用，將使用時間鎖定至使用日上午十點");
                            jtem.SaleTime = "10:00";
                        }
                    }

                    if (!errorData)
                    {
                        Month_GroupTickets.AddRange(item);
                    }
                    else
                    {
                        MyCommond.WriteLog(ThisReceive, $"資料錯誤。");
                        var ss = MessageBox.Show($"資料有誤。(路徑:{GroupTicketPath})\n" +
                            $"請確認後再點擊確認重新載入這筆資料。\n" +
                            $" 是 :重新執行這個檔案。\n" +
                            $" 否 :跳過這個檔案。\n" +
                            $"取消:結束執行。",
                            "錯誤", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                        if (ss == DialogResult.Yes)
                        {
                            goto LoadData;
                        }
                        else if (ss == DialogResult.No)
                        {
                            continue;
                        }
                        else if (ss == DialogResult.Cancel)
                        {
                            return false;
                        }
                    }
                }
            }

            #endregion
            return true;
        }

        #region 月運量統計表(停用)

        private void Calculator_Month_AllRideList(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); Data_Afc = ArrayToExport(Calculator_1()); }

                Data_MobilePay = null;

                Data_OwnTicket = null;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }
        }

        #endregion

        #region 每月起訖總表

        private void Calculator_Month_OriginDestination(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                var station_ = MainServer.globalCommond.Stock_Lrt_Station.ToList();
                int stationCount = station_.Count;
                ValueCount += 1;

                MyCommond.WriteLog(ThisReceive, $"開始相關電票資料擷取");
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); Data_Afc = ArrayToExport(Calculator_1()); }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_Afc, "Data_Afc");
                ValueCount = 40;

                MyCommond.WriteLog(ThisReceive, $"開始相關電支資料擷取");
                Data_MobilePay = null;
                foreach (var item in Month_MobilePayData)
                {
                    if (item.Volume == MobilePayVolum.InTime) /*------*/ { continue; }
                    int EntryStation = station_.FindIndex(x => x.StationName == item.EntryStation);
                    int ExitStation_ = station_.FindIndex(x => x.StationName == item.ExitStation);
                    if (!(EntryStation > -1 && ExitStation_ > -1)) /*-*/ { continue; }
                    Data_Afc[EntryStation, ExitStation_]++;
                }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_MobilePay, "Data_MobilePay");
                ValueCount = 60;

                MyCommond.WriteLog(ThisReceive, $"開始相關團票資料擷取");
                Data_OwnTicket = null;
                foreach (var item in Month_GroupTickets)
                {
                    int EntryStation = station_.FindIndex(x => x.CodeName == item.Entry);
                    int ExitStation_ = station_.FindIndex(x => x.CodeName == item.Exit);
                    Data_Afc[EntryStation, ExitStation_] += Convert.ToInt32(item.People);
                }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_OwnTicket, "Data_OwnTicket");
                ValueCount = 80;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }
        }

        #endregion

        #region 各家票卡運量月報(車站)

        private void Calculator_Month_ElectronicTicket_Station(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                MyCommond.WriteLog(ThisReceive, $"開始相關電票資料擷取");
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); Data_Afc = ArrayToExport(Calculator_1()); }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_Afc, "Data_Afc");
                ValueCount = 40;

                MyCommond.WriteLog(ThisReceive, $"開始相關電支資料擷取");
                Data_MobilePay = Volume_Mobile_5_Line(Month_MobilePayData, ReportSwitch_1.Station);
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_MobilePay, "Data_MobilePay");
                ValueCount = 60;

                MyCommond.WriteLog(ThisReceive, $"無須填入團體票");
                Data_OwnTicket = null;
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_OwnTicket, "Data_OwnTicket");
                ValueCount = 80;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }
        }

        #endregion

        #region 各家票卡運量月報(天期)

        private void Calculator_Month_ElectronicTicket_Day(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                MyCommond.WriteLog(ThisReceive, $"開始相關電票資料擷取");
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); Data_Afc = ArrayToExport(Calculator_1()); }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_Afc, "Data_Afc");
                ValueCount = 40;

                MyCommond.WriteLog(ThisReceive, $"開始相關電支資料擷取");
                Data_MobilePay = Volume_Mobile_5_Line(Month_MobilePayData, ReportSwitch_1.Day);
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_MobilePay, "Data_MobilePay");
                ValueCount = 60;

                MyCommond.WriteLog(ThisReceive, $"開始相關團票資料擷取");
                Data_OwnTicket = null;
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_OwnTicket, "Data_OwnTicket");
                ValueCount = 80;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }
        }

        #endregion

        #region 自有票卡運量月報(車站)

        private void Calculator_Month_OwnTicketVolume_Station(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); Data_Afc = ArrayToExport(Calculator_1()); }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_Afc, "Data_Afc");
                ValueCount = 40;

                MyCommond.WriteLog(ThisReceive, $"開始相關電支資料擷取");
                Data_MobilePay = null;
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_MobilePay, "Data_MobilePay");
                ValueCount = 60;

                MyCommond.WriteLog(ThisReceive, $"開始相關團票資料擷取");
                Data_OwnTicket = new int[MainServer.globalCommond.Stock_Lrt_Station.Count, 9];
                if (MainServer.globalCommond.Stock_Lrt_PaoStation.Count == 0) { MyCommond.WriteLog(ThisReceive, "PaoData 異常"); }
                foreach (var item in MainServer.globalCommond.Stock_Lrt_PaoStation)
                {
                    var GroupPao = MainServer.paoTicketSaleLists.FindAll(x => x.Station == item.CodeName).ToList();

                    var StationIndex = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.CodeName == item.CodeName);

                    var SOneTicket = GroupPao.Sum(x => Convert.ToInt32(x.Ticket20)) + GroupPao.Sum(x => Convert.ToInt32(x.Ticket25)) + GroupPao.Sum(x => Convert.ToInt32(x.Ticket30));
                    var SOneTicketCon = GroupPao.Sum(x => Convert.ToInt32(x.Ticket10)) + GroupPao.Sum(x => Convert.ToInt32(x.Ticket12)) + GroupPao.Sum(x => Convert.ToInt32(x.Ticket15));
                    var SOneTicketBike = GroupPao.Sum(x => Convert.ToInt32(x.TicketBike));
                    var OneDay = GroupPao.Sum(x => Convert.ToInt32(x.Ticket_OneDay));
                    var OneDayLanwair = GroupPao.Sum(x => Convert.ToInt32(x.Lanwair1) + Convert.ToInt32(x.Lanwair2));
                    var GroupCount = GroupPao.Sum(x => Convert.ToInt32(x.GroupCount));
                    var GroupPeople = GroupPao.Sum(x => Convert.ToInt32(x.GroupPeople));

                    Data_OwnTicket[StationIndex, 0] += SOneTicket;
                    Data_OwnTicket[StationIndex, 1] += SOneTicketCon;
                    Data_OwnTicket[StationIndex, 2] += SOneTicketBike;
                    Data_OwnTicket[StationIndex, 3] += OneDay;
                    Data_OwnTicket[StationIndex, 4] += OneDay * 6;
                    Data_OwnTicket[StationIndex, 5] += OneDayLanwair;
                    Data_OwnTicket[StationIndex, 6] += OneDayLanwair * 6;
                    Data_OwnTicket[StationIndex, 7] += GroupCount;
                    Data_OwnTicket[StationIndex, 8] += GroupPeople;

                }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_OwnTicket, "Data_OwnTicket");
                ValueCount = 80;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }
        }

        #endregion

        #region 自有票卡運量月報(天期)

        private void Calculator_Month_OwnTicketVolume_Day(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                MyCommond.WriteLog(ThisReceive, $"開始相關電票資料擷取");
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"開始相關電票資料擷取"); Data_Afc = ArrayToExport(Calculator_1()); }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_Afc, "Data_Afc");
                ValueCount = 40;

                MyCommond.WriteLog(ThisReceive, $"開始相關電支資料擷取");
                Data_MobilePay = null;
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_MobilePay, "Data_MobilePay");
                ValueCount = 60;

                MyCommond.WriteLog(ThisReceive, $"開始相關團票資料擷取");
                Data_OwnTicket = new int[Data_Afc.GetLength(0), 9];
                if (MainServer.globalCommond.Stock_Lrt_PaoStation.Count == 0) { MyCommond.WriteLog(ThisReceive, "PaoData 異常"); }
                foreach (var item in MainServer.paoTicketSaleLists)
                {
                    //var mYear = item.Operation.Substring(0,4);
                    //var mMonth = item.Operation.Substring(4,2);
                    var mDay = item.Operation.Substring(6, 2);
                    var RowIndex = Convert.ToInt32(mDay) - 1;

                    var SOneTicket = Convert.ToInt32(item.Ticket20) + Convert.ToInt32(item.Ticket25) + Convert.ToInt32(item.Ticket30);
                    var SOneTicketCon = Convert.ToInt32(item.Ticket10) + Convert.ToInt32(item.Ticket12) + Convert.ToInt32(item.Ticket15);
                    var SOneTicketBike = Convert.ToInt32(item.TicketBike);
                    var OneDay = Convert.ToInt32(item.Ticket_OneDay);
                    var OneDayLanwair = Convert.ToInt32(item.Lanwair1) + Convert.ToInt32(item.Lanwair2);
                    var GroupCounts = Convert.ToInt32(item.GroupCount);
                    var GroupPeople = Convert.ToInt32(item.GroupPeople);

                    Data_OwnTicket[RowIndex, 0] += SOneTicket;
                    Data_OwnTicket[RowIndex, 1] += SOneTicketCon;
                    Data_OwnTicket[RowIndex, 2] += SOneTicketBike;
                    Data_OwnTicket[RowIndex, 3] += OneDay;
                    Data_OwnTicket[RowIndex, 4] += OneDay * 6;
                    Data_OwnTicket[RowIndex, 5] += OneDayLanwair;
                    Data_OwnTicket[RowIndex, 6] += OneDayLanwair * 6;
                    Data_OwnTicket[RowIndex, 7] += GroupCounts;
                    Data_OwnTicket[RowIndex, 8] += GroupPeople;
                }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_OwnTicket, "Data_OwnTicket");
                ValueCount = 80;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }
        }

        #endregion

        #region 營收月報表(車站)

        private void Calculator_Month_TrafficAmount_Station(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                MyCommond.WriteLog(ThisReceive, $"開始相關電票資料擷取");
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); Data_Afc = ArrayToExport(Calculator_1()); }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_Afc, "Data_Afc");
                ValueCount = 40;

                MyCommond.WriteLog(ThisReceive, $"開始相關電支資料擷取");
                Data_MobilePay = Amount_Mobile_2_Line(Month_MobilePayData, ReportSwitch_1.Station);
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_MobilePay, "Data_MobilePay");
                ValueCount = 60;

                MyCommond.WriteLog(ThisReceive, $"開始相關人工票資料擷取");
                Data_OwnTicket = new int[MainServer.globalCommond.Stock_Lrt_Station.Count, 12];
                if (MainServer.globalCommond.Stock_Lrt_PaoStation.Count == 0) { MyCommond.WriteLog(ThisReceive, "PaoData 異常"); }
                foreach (var item in MainServer.globalCommond.Stock_Lrt_PaoStation)
                {
                    var GroupPao = MainServer.paoTicketSaleLists.FindAll(x => x.Station == item.CodeName).ToList();

                    var StationIndex = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.CodeName == item.CodeName);

                    Data_OwnTicket[StationIndex, 0] = GroupPao.Sum(x => Convert.ToInt32(x.Ticket20)) * 20 + GroupPao.Sum(x => Convert.ToInt32(x.Ticket25)) * 25 + GroupPao.Sum(x => Convert.ToInt32(x.Ticket30)) * 30;
                    Data_OwnTicket[StationIndex, 1] = GroupPao.Sum(x => Convert.ToInt32(x.Ticket10)) * 10 + GroupPao.Sum(x => Convert.ToInt32(x.Ticket12)) * 12 + GroupPao.Sum(x => Convert.ToInt32(x.Ticket15)) * 15;
                    Data_OwnTicket[StationIndex, 2] = GroupPao.Sum(x => Convert.ToInt32(x.TicketBike)) * 50;
                    Data_OwnTicket[StationIndex, 3] = GroupPao.Sum(x => Convert.ToInt32(x.Ticket_OneDay)) * 50;
                    Data_OwnTicket[StationIndex, 4] = GroupPao.Sum(x => Convert.ToInt32(x.Lanwair1)) * 50;
                    Data_OwnTicket[StationIndex, 5] = GroupPao.Sum(x => Convert.ToInt32(x.GroupAmt));
                    Data_OwnTicket[StationIndex, 6] = GroupPao.Sum(x => Convert.ToInt32(x.OpenTrafficSetTicket)) * 450;
                    Data_OwnTicket[StationIndex, 7] = GroupPao.Sum(x => Convert.ToInt32(x.RepairTicket));
                    Data_OwnTicket[StationIndex, 8] = GroupPao.Sum(x => Convert.ToInt32(x.AfcTakeOut));
                    Data_OwnTicket[StationIndex, 9] = GroupPao.Sum(x => Convert.ToInt32(x.VavmRefund));
                    Data_OwnTicket[StationIndex, 10] = GroupPao.Sum(x => Convert.ToInt32(x.SaleCardRefund));
                    Data_OwnTicket[StationIndex, 11] = GroupPao.Sum(x => Convert.ToInt32(x.PvRefund));

                }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_OwnTicket, "Data_OwnTicket");
                ValueCount = 80;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }
        }

        #endregion

        #region 營收月報表(天期)

        private void Calculator_Month_TrafficAmount_Day(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                MyCommond.WriteLog(ThisReceive, $"開始相關電票資料擷取");
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); Data_Afc = ArrayToExport(Calculator_1()); }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_Afc, "Data_Afc");
                ValueCount = 40;

                MyCommond.WriteLog(ThisReceive, $"開始相關電支資料擷取");
                Data_MobilePay = Amount_Mobile_2_Line(Month_MobilePayData, ReportSwitch_1.Day);
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_MobilePay, "Data_MobilePay");
                ValueCount = 60;

                MyCommond.WriteLog(ThisReceive, $"開始相關人工票資料擷取");
                Data_OwnTicket = new int[Data_Afc.GetLength(0), 12];
                if (MainServer.globalCommond.Stock_Lrt_PaoStation.Count == 0) { MyCommond.WriteLog(ThisReceive, "PaoData 異常"); }
                foreach (var item in MainServer.paoTicketSaleLists)
                {
                    //var mYear = item.Operation.Substring(0,4);
                    //var mMonth = item.Operation.Substring(4,2);
                    var mDay = item.Operation.Substring(6, 2);
                    var RowIndex = Convert.ToInt32(mDay) - 1;

                    var SOneTicket = Convert.ToInt32(item.Ticket20) * 20 + Convert.ToInt32(item.Ticket25) * 25 + Convert.ToInt32(item.Ticket30) * 30;
                    var SOneTicketCon = Convert.ToInt32(item.Ticket10) * 10 + Convert.ToInt32(item.Ticket12) * 12 + Convert.ToInt32(item.Ticket15) * 15;
                    var SOneTicketBike = Convert.ToInt32(item.TicketBike) * 50;
                    var OneDay = Convert.ToInt32(item.Ticket_OneDay) * 50;
                    var GroupTicket = Convert.ToInt32(item.GroupAmt);
                    var OpenTicket = Convert.ToInt32(item.OpenTrafficSetTicket) * 450;
                    var Repair = Convert.ToInt32(item.RepairTicket);
                    var TakeOut = Convert.ToInt32(item.AfcTakeOut);
                    var ReVavm = Convert.ToInt32(item.VavmRefund);
                    var ReCard = Convert.ToInt32(item.SaleCardRefund);
                    var RePv = Convert.ToInt32(item.PvRefund);

                    Data_OwnTicket[RowIndex, 0] += SOneTicket;
                    Data_OwnTicket[RowIndex, 1] += SOneTicketCon;
                    Data_OwnTicket[RowIndex, 2] += SOneTicketBike;
                    Data_OwnTicket[RowIndex, 3] += OneDay;
                    Data_OwnTicket[RowIndex, 4] += 0;
                    Data_OwnTicket[RowIndex, 5] += GroupTicket;
                    Data_OwnTicket[RowIndex, 6] += OpenTicket;
                    Data_OwnTicket[RowIndex, 7] += Repair;
                    Data_OwnTicket[RowIndex, 8] += TakeOut;
                    Data_OwnTicket[RowIndex, 9] += ReVavm;
                    Data_OwnTicket[RowIndex, 10] += ReCard;
                    Data_OwnTicket[RowIndex, 11] += RePv;
                }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_OwnTicket, "Data_OwnTicket");
                ValueCount = 80;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }
        }

        #endregion

        #region 設備營收月報表(車站)

        private void Calculator_Month_EquipAmount_Station(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                MyCommond.WriteLog(ThisReceive, $"開始相關電票資料擷取");
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); Data_Afc = ArrayToExport(Calculator_1()); }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_Afc, "Data_Afc");
                ValueCount = 40;

                MyCommond.WriteLog(ThisReceive, $"開始相關電支資料擷取");
                Data_MobilePay = Amount_10_Line(Month_MobilePayData, ReportSwitch_1.Station);
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_MobilePay, "Data_MobilePay");
                ValueCount = 60;

                MyCommond.WriteLog(ThisReceive, $"開始相關團票資料擷取");
                Data_OwnTicket = null;
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_OwnTicket, "Data_OwnTicket");
                ValueCount = 20;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }
        }

        #endregion

        #region 設備營收月報表(天期)

        private void Calculator_Month_EquipAmount_Day(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            ThisMethode = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {
                MyCommond.WriteLog(ThisReceive, $"開始相關電票資料擷取");
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); Data_Afc = ArrayToExport(Calculator_1()); }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_Afc, "Data_Afc");
                ValueCount = 40;

                MyCommond.WriteLog(ThisReceive, $"開始相關電支資料擷取");
                Data_MobilePay = Amount_10_Line(Month_MobilePayData, ReportSwitch_1.Day);
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_MobilePay, "Data_MobilePay");
                ValueCount = 60;

                MyCommond.WriteLog(ThisReceive, $"開始相關團票資料擷取");
                Data_OwnTicket = null;
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_OwnTicket, "Data_OwnTicket");
                ValueCount = 80;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }
        }

        #endregion
        
        #region 營運資料統計月報_運量

        private void Calculator_Month_Volume(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            ThisMethode = StatusEnum.Calculating;

            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {

                #region 電票

                MyCommond.WriteLog(ThisReceive, $"取得電子票證資料");
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); Data_Afc = ArrayToExport(Calculator_1()); }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_Afc, "Data_Afc");

                ValueCount = 40;

                #endregion

                #region 電支

                MyCommond.WriteLog(ThisReceive, $"開始相關電支資料擷取");
                Data_MobilePay = Mobile_2(Month_MobilePayData, ReportSwitch_1.Day, ReportSwitch_2.Volumn);
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_MobilePay, "Data_MobilePay");
                ValueCount = 60;

                #endregion

                #region 人工

                MyCommond.WriteLog(ThisReceive, $"取得Pao收入、支出資料");
                Data_OwnTicket = PaoCal(MainServer.paoTicketSaleLists, Month_GroupTickets, ReportSwitch_1.Day, ReportSwitch_2.Volumn);
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_OwnTicket, "Data_OwnTicket");
                ValueCount = 80;

                #endregion

            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }

            ValueCount = 80;
        }

        #endregion

        #region 營運資料統計月報_營收

        private void Calculator_Month_Amount(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            ValueCount = 20;
            ThisMethode = StatusEnum.Calculating;

            MyCommond.WriteLog(ThisReceive, $"開始計算{ThisMethode.ToDescription()}");

            try
            {

                #region 電票

                MyCommond.WriteLog(ThisReceive, $"取得電子票證資料");
                if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); Data_Afc = ArrayToExport(Calculator_1()); }
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_Afc, "Data_Afc");
                ValueCount = 30;

                #endregion

                #region 電支

                MyCommond.WriteLog(ThisReceive, $"開始相關電支資料擷取");
                Data_MobilePay = Mobile_2(Month_MobilePayData, ReportSwitch_1.Day, ReportSwitch_2.Amount);
                //MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data_MobilePay, "Data_MobilePay");
                ValueCount = 40;

                #endregion

                #region 人工

                MyCommond.WriteLog(ThisReceive, $"取得Pao收入、支出資料");
                Data_OwnTicket = PaoCal(MainServer.paoTicketSaleLists, Month_GroupTickets, ReportSwitch_1.Day, ReportSwitch_2.Amount);
                ValueCount = 50;

                #endregion

                #region OD

                MyCommond.WriteLog(ThisReceive, $"運價票差 - 開始");

                int DateCount = (MainServer.OperationEndDate - MainServer.OperationStartDate).Days;

                int[,] ArrayDate_SubsidyAmt = new int[DateCount, 2];

                var Station_CodeNumber = (from p in MainServer.globalCommond.Stock_Lrt_Station select Convert.ToByte(p.CodeNumber)).ToList();
                var SqlFileName_Od = MainServer.globalCommond.Stock_SqlList.Find(x => x.UseNameCh == LrtTxnClientSwitch.Day_OriginDestination.ToDescription()).FileName1;
                var Sqlstr_Od = File.ReadAllText(MyCommond.Path_Program + MyCreat.Path_Sql + SqlFileName_Od + ".txt");

                for (int z = 0; z < DateCount; z++)
                {
                    DateTime dt = MainServer.OperationStartDate.AddDays(z);

                    MyCommond.WriteLog(ThisReceive, $"取得 {dt.ToString("yyyy/MM/dd")} 運價票差");

                    int SubsidyAmt = 0;

                    #region 電支(OD)

                    MyCommond.WriteLog(ThisReceive, $"計算電支運價票差");

                    var item = (from p in Month_MobilePayData where p.ExitTime < dt.AddDays(1) select p).ToList();

                    foreach (var jtem in item)
                    {
                        string Station_I = "";
                        string Station_O = "";
                        string TicketSubsidy = "";

                        try
                        {
                            Station_I = MainServer.globalCommond.Stock_Lrt_Station.Find(x => x.StationName == jtem.EntryStation).CodeName;
                            Station_O = MainServer.globalCommond.Stock_Lrt_Station.Find(x => x.StationName == jtem.ExitStation).CodeName;
                            TicketSubsidy = MainServer.globalCommond.Stock_Lrt_Subsidy.Find(x =>
                            x.Operation_Area == MainServer.globalCommond.OperationCode &&
                            x.StartStation == Station_I &&
                            x.EndStation == Station_O).OdAmt;

                            SubsidyAmt += Convert.ToInt32(TicketSubsidy);
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error\n{ex.Message}");
                        }
                        finally
                        {
                            MyCommond.WriteLog(ThisReceive, $"進站:{Station_I,4}, 出站:{Station_O,4}, 補貼:{TicketSubsidy,4}, 比數:{1,8}, 累計補貼金額:{SubsidyAmt}");
                        }
                    }

                    #endregion

                    #region 電票(OD)

                    MyCommond.WriteLog(ThisReceive, $"取得電票OD");

                    var SqlCommand_Od = string.Format(Sqlstr_Od, Station_CodeNumber.Min(), Station_CodeNumber.Max(), dt.ToString("yyyy/MM/dd 05:00:00"), dt.AddDays(1).ToString("yyyy/MM/dd 05:00:00"));

                    string[,] SqlOd = inputZ(new string[Station_CodeNumber.Count, Station_CodeNumber.Count + 1]);
                    if (MyCommond.CheckIP(ThisReceive, MainServer.globalCommond.SqlServers.IP)) { MyCommond.WriteLog(ThisReceive, $"a"); SqlOd = SQLTextSelect(SqlCommand_Od); }

                    MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), SqlOd, "SqlOd");

                    MyCommond.WriteLog(ThisReceive, $"計算電票運價票差");

                    for (int i = 0; i < SqlOd.GetLength(0); i++)
                    {
                        for (int j = 1; j < SqlOd.GetLength(1); j++)
                        {
                            string Station_I = "";
                            string Station_O = "";
                            string TicketSubsidy = "";
                            int DataCount = 0;
                            try
                            {
                                Station_I = MainServer.globalCommond.Stock_Lrt_Station[i].CodeName;
                                Station_O = MainServer.globalCommond.Stock_Lrt_Station[j - 1].CodeName;
                                TicketSubsidy = MainServer.globalCommond.Stock_Lrt_Subsidy.Find(x => 
                                    x.Operation_Area == MainServer.globalCommond.OperationCode && 
                                    x.StartStation == Station_I && 
                                    x.EndStation == Station_O).OdAmt;

                                DataCount = (SqlOd[i, j] != null || SqlOd[i, j] != "") ? Convert.ToInt32(SqlOd[i, j]) : 0;
                                if (TicketSubsidy != null && Convert.ToInt32(TicketSubsidy) > 0) { SubsidyAmt += (Convert.ToInt32(TicketSubsidy) * DataCount); }
                                // MyCommond.WriteLog(ThisReceive, $"進站:{Station_I,4}, 出站:{Station_O,4}, 補貼:{TicketSubsidy,4}, 比數:{DataCount,8}, 累計補貼金額:{SubsidyAmt}");
                            }
                            catch (Exception ex)
                            {
                                MyCommond.WriteLog(ThisReceive, $"Error\n{ex.Message}");
                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"進站:{Station_I,4}, 出站:{Station_O,4}, 補貼:{TicketSubsidy,4}, 比數:{DataCount,8}, 累計補貼金額:{SubsidyAmt}");
                            }
                        }
                    }

                    #endregion

                    Data_Afc[dt.Day - 1, Data_Afc.GetLength(1) - 1] += SubsidyAmt;

                }
                ValueCount = 70;

                #endregion

            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
            finally
            {

            }

            ValueCount = 80;
        }

        #endregion

        #endregion

        #endregion

        #region 處理

        private void ListDataCount(List<MobilePay> Data)
        {
            MyCommond.WriteLog(ThisReceive, $"資料總筆數:{Data.Count}");
            for (var dt = MainServer.OperationStartDate; dt < MainServer.OperationEndDate; dt = dt.AddDays(1))
            {
                var tempData = Data.FindAll(x => dt < x.ExitTime && x.ExitTime < dt.AddDays(1)).ToList();
                var tempRealPay = tempData.Sum(x => x.RealPay);
                var tempShoudPay = tempData.Sum(x => x.ShouldPay);
                var tempFine = tempData.Sum(x => x.Fine);
                MyCommond.WriteLog(ThisReceive, $"" +
                    $"  日期  : {dt:yyyy/MM/dd}\n" +
                    $"應有筆數: {tempData.Count,5} 筆\n" +
                    $"應付車資: {tempShoudPay,10} 元\n" +
                    $"  罰款  : {tempFine,10} 元\n" +
                    $"實收金額: {tempRealPay,10} 元");
            }

            if (!true)
            {
                MyCommond.ExportData2<MobilePay>(ThisReceive, new ExportConfig()
                {
                    Path1 = MyCommond.Path_Program + MyCommond.Path_Report,
                    FirstName = MainServer.OperationStartDate.ToString("yyyyMMdd"),
                    FileName = "MobilePayData",
                    SubName = "進程式判斷的資料",
                    OpLine = MainServer.globalCommond.OperationCode_String,
                    LastName = "xlsx",
                }, new MobilePay().ExportTitle_zhTW(), Data, Encoding.UTF8);
            }

        }

        private void FileNameExample(ExportConfig ec)
        {
            var mYear = ((Convert.ToInt32(ec.DTime.ToString("yyyy")) > 1911) ? (Convert.ToInt32(ec.DTime.ToString("yyyy")) - 1911) : Convert.ToInt32(ec.DTime.ToString("yyyy"))).ToString();
            var mMonth = ec.DTime.ToString("MM");
            var mDay = ec.DTime.ToString("dd");

            string filePath = ec.Path1;
            string fileName1 = $"{mYear}{mMonth}{mDay}";
            if (MethodeName >= LrtTxnClientSwitch.Month_AllRideList) { fileName1 = $"{mYear}{mMonth}"; }
            string fileName2 = ec.FirstName;
            string fileName3 = ec.FileName;
            string fileName4 = ec.SubName;
            ValueCount += 1;

            string FullFileName = ((fileName1.Length > 0) ? $"{fileName1}_" : "") +
                ((fileName2.Length > 0) ? $"{fileName2}_" : "") +
                ((fileName3.Length > 0) ? $"{fileName3}_" : "") +
                ((fileName4.Length > 0) ? $"{fileName4}_" : "") +
                ec.OpLine;

            ExportPath = filePath;

            ExportFileName = MyCommond.CheckAndReName(ExportPath, FullFileName, ec.LastName);
            MyCommond.WriteLog(ThisReceive, $"變更輸出位置:{ExportPath}");
            MyCommond.WriteLog(ThisReceive, $"變更輸出檔名:{ExportFileName}");
            ValueCount += 1;
        }

        private string[,] Calculator_1()
        {
            MyCommond.WriteLog(ThisReceive, $"開始計算 - {MethodeName.ToDescription()}");
            ValueCount = 40;

            var SqlFileName = MainServer.globalCommond.Stock_SqlList.Find(x => x.UseNameCh == MethodeName.ToDescription()).FileName1;
            var Sqlstr = File.ReadAllText(MyCommond.Path_Program + MyCreat.Path_Sql + SqlFileName + ".txt");
            var Station = (from p in MainServer.globalCommond.Stock_Lrt_Station select Convert.ToByte(p.CodeNumber)).ToList();
            var OperationData_Start = OperationDate.ToString("yyyy-MM-dd 05:00:00");
            var OperationData_End = OperationDate.AddDays(1).ToString("yyyy-MM-dd 05:00:00");
            if (MethodeName >= LrtTxnClientSwitch.Month_AllRideList)
            {
                OperationData_Start = MainServer.OperationStartDate.ToString("yyyy-MM-dd 05:00:00");
                OperationData_End = MainServer.OperationEndDate.ToString("yyyy-MM-dd 05:00:00");
            }
            var SqlCommand = string.Format(Sqlstr, Station.Min(), Station.Max(), OperationData_Start, OperationData_End);
            ValueCount = 45;

            string[] Nothing = null;
            var item = SQLTextSelect(SqlCommand);
            ValueCount = 50;

            return item;
        }

        private void Calculator_1(ref int[,] Data_Afc, ref int[,] Data_MobilePay, ref int[,] Data_OwnTicket)
        {
            MyCommond.WriteLog(ThisReceive, $"開始計算 - {MethodeName.ToDescription()}");
            ValueCount = 40;

            var SqlFileName = MainServer.globalCommond.Stock_SqlList.Find(x => x.UseNameCh == MethodeName.ToDescription()).FileName1;
            var Sqlstr = File.ReadAllText(MyCommond.Path_Program + MyCreat.Path_Sql + SqlFileName + ".txt");
            var Station = (from p in MainServer.globalCommond.Stock_Lrt_Station select Convert.ToByte(p.CodeNumber)).ToList();
            var SqlCommand = string.Format(Sqlstr, Station.Min(), Station.Max(), OperationDate.ToString("yyyy-MM-dd 05:00:00"), OperationDate.AddDays(1).ToString("yyyy-MM-dd 05:00:00"));

            ValueCount = 45;

            Data_Afc = ArrayToExport(SQLTextSelect(SqlCommand));

            ValueCount = 50;

        }

        /** private bool CheckIP()
        {
            MyCommond.WriteLog(ThisReceive, $"check..");
            string strHostName = string.Empty;
            IPHostEntry ipEntry = Dns.GetHostEntry(Dns.GetHostName());
            IPAddress[] addr = ipEntry.AddressList;

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

                if (cc == 3) { return true; }
            }
            return false;
        }
        */
        private string[,] SQLTextSelect(string commandText)
        {
            string sql_command_string = @"Data Source='{0}';Initial Catalog={1};Persist Security Info=True;User ID='{2}';Password='{3}'; Timeout=60 ";
            string sql_command = string.Format(sql_command_string,
                MainServer.globalCommond.SqlServers.IP,
                MainServer.globalCommond.SqlServers.Calalog,
                MainServer.globalCommond.SqlServers.ID,
                MainServer.globalCommond.SqlServers.Password);

            DataTable dt = new DataTable();
            int Reconnect = 0;

            MyCommond.WriteLog(ThisReceive, $"準備資料庫連線");
        StartConnect:
            using (SqlConnection con = new SqlConnection(sql_command))
            {
                using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand())
                {
                    cmd.CommandText = commandText;
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = con;
                    cmd.CommandTimeout = (((int)MethodeName > 8) ? 600 : 60);
                    MyCommond.WriteLog(ThisReceive, $"寫入資料庫語法\n{commandText}");

                    try
                    {
                        con.Open();
                        SqlDataReader reader = cmd.ExecuteReader();

                        ValueCount += 1;

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
                            MyCommond.WriteLog(ThisReceive, $"ERROR!\n{MyCommond.ExportErrorMessageToLog(ex, MethodeName.ToDescription()).Message}");
                        }
                    }
                    finally
                    {
                        con.Close();
                    }
                }
            }

            if (Reconnect >= 3) return null;

            if (dt == null || dt.Rows.Count <= 0)
            {
                MyCommond.WriteLog(ThisReceive, $"重新查詢");
                Reconnect++;
                goto StartConnect;
            }

            ValueCount += 1;

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

            MyCommond.WriteLog(ThisReceive, $"資料撈取結束");
            return vs;
        }

        private string[,] SQLTextSelect(string commandText, ref string[] ResponTitle)
        {
            string sql_command_string = @"Data Source='{0}';Initial Catalog={1};Persist Security Info=True;User ID='{2}';Password='{3}'; Timeout=60 ";
            string sql_command = string.Format(sql_command_string,
                MainServer.globalCommond.SqlServers.IP,
                MainServer.globalCommond.SqlServers.Calalog,
                MainServer.globalCommond.SqlServers.ID,
                MainServer.globalCommond.SqlServers.Password);

            DataTable dt = new DataTable();

            MyCommond.WriteLog(ThisReceive, $"準備資料庫連線");
            using (SqlConnection con = new SqlConnection(sql_command))
            {
                using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand())
                {
                    cmd.CommandText = commandText;
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = con;
                    cmd.CommandTimeout = (((int)MethodeName > 8) ? 600 : 60);
                    MyCommond.WriteLog(ThisReceive, $"寫入資料庫語法\n{commandText}");

                    try
                    {
                        con.Open();
                        SqlDataReader reader = cmd.ExecuteReader();

                        ValueCount += 1;

                        MyCommond.WriteLog(ThisReceive, $"載入資料庫取得的資料");
                        dt.Load(reader);
                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $"ERROR!\n{MyCommond.ExportErrorMessageToLog(ex, MethodeName.ToDescription()).Message}");
                    }
                    finally
                    {
                        con.Close();
                    }
                }
            }

            ValueCount += 1;

            MyCommond.WriteLog(ThisReceive, $"將資料轉換成數字陣列");
            MyCommond.WriteLog(ThisReceive, $"行:{dt.Rows.Count,3}, 欄:{dt.Columns.Count,3}");
            string[,] vs = new string[dt.Rows.Count, dt.Columns.Count];

            var dr = dt.CreateDataReader();
            var dr_c = dr.FieldCount;
            ResponTitle = new string[dr_c];
            for (int i = 0; i < dr_c; i++)
            {
                MyCommond.WriteLog(ThisReceive, $"dr.GetName({i}) = {dr.GetName(i)}");
                ResponTitle[i] = dr.GetName(i);
            }

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    vs[i, j] = Convert.ToString(dt.Rows[i][j]);
                }
            }

            MyCommond.WriteLog(ThisReceive, $"資料撈取結束");
            return vs;
        }

        private DataTableReader GetSqlData(string commandText)
        {
            string sql_command_string = @"Data Source='{0}';Initial Catalog={1};Persist Security Info=True;User ID='{2}';Password='{3}'; Timeout=60 ";
            string sql_command = string.Format(sql_command_string,
                MainServer.globalCommond.SqlServers.IP,
                MainServer.globalCommond.SqlServers.Calalog,
                MainServer.globalCommond.SqlServers.ID,
                MainServer.globalCommond.SqlServers.Password);

            DataTable dt = new DataTable();
            DataTableReader dr = null;

            MyCommond.WriteLog(ThisReceive, $"準備資料庫連線");
            using (SqlConnection con = new SqlConnection(sql_command))
            {
                using (System.Data.SqlClient.SqlCommand cmd = new System.Data.SqlClient.SqlCommand())
                {
                    cmd.CommandText = commandText;
                    cmd.CommandType = CommandType.Text;
                    cmd.Connection = con;
                    cmd.CommandTimeout = (((int)MethodeName > 8) ? 600 : 60);
                    MyCommond.WriteLog(ThisReceive, $"寫入資料庫語法\n{commandText}");

                    try
                    {
                        con.Open();
                        SqlDataReader reader = cmd.ExecuteReader();

                        ValueCount += 1;

                        MyCommond.WriteLog(ThisReceive, $"載入資料庫取得的資料");
                        dt.Load(reader);
                        dr = dt.CreateDataReader();
                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $"ERROR!\n{MyCommond.ExportErrorMessageToLog(ex, MethodeName.ToDescription()).Message}");
                    }
                    finally
                    {
                        con.Close();
                    }
                }
            }

            ValueCount += 1;

            MyCommond.WriteLog(ThisReceive, $"行:{dt.Rows.Count,3}, 欄:{dt.Columns.Count,3}");

            var dr_c = dr.FieldCount;
            for (int i = 0; i < dr_c; i++)
            {
                MyCommond.WriteLog(ThisReceive, $"dr.GetName({i}) = {dr.GetName(i)}");
            }

            ValueCount += 1;

            MyCommond.WriteLog(ThisReceive, $"資料撈取結束");
            return dr;
        }

        private List<string[]> DataToList(DataTableReader TData, ref string[] ResponTitle)
        {
            MyCommond.WriteLog(ThisReceive, $"資料表轉換為清單 - 開始");
            List<string[]> Result = new List<string[]>();
            int ListRowCount = 0;

            bool writeTitle = ResponTitle != null;
            MyCommond.WriteLog(ThisReceive, $"writeTitle = {writeTitle}");

            int DataLenght = TData.FieldCount;
            MyCommond.WriteLog(ThisReceive, $"取得欄位數量: {DataLenght}");

            while (TData.Read())
            {
                if (!writeTitle)
                {
                    MyCommond.WriteLog(ThisReceive, $"寫入欄位名稱");
                    ResponTitle = new string[DataLenght];
                    for (int i = 0; i < DataLenght; i++)
                    {
                        ResponTitle[i] = TData.GetName(i);
                        MyCommond.WriteLog(ThisReceive, $"ResponTitle[{i,4}] = {ResponTitle[i]}");
                    }
                    writeTitle = ResponTitle != null;
                    MyCommond.WriteLog(ThisReceive, $"writeTitle = {writeTitle}");
                }

                string[] tempData = new string[DataLenght];

                MyCommond.WriteLog(ThisReceive, $"寫入資料");
                for (int i = 0; i < DataLenght; i++)
                {
                    MyCommond.WriteLog(ThisReceive, $"{ListRowCount,4}.tempData[{i,4}] = {TData.GetString(i)}");
                    tempData[i] = TData.GetString(i);
                }

                Result.Add(tempData);
                ListRowCount++;
            }

            MyCommond.WriteLog(ThisReceive, $"資料表轉換為清單 - 結束");
            return Result;
        }

        private int[,] Volume_CreatColumn(string CardKind, string MP_Volumn, MobileEnum IO_Add)
        {
            var Station_Count = MainServer.globalCommond.Stock_Lrt_Station.Count;
            var CardKindList = MainServer.globalCommond.Stock_MobileCardType.FindAll(x => x.CardType == CardKind).ToList();
            var R_Data = (from p in Day_MobilePayData where CardKindList.FindIndex(x => x.CardSubType == p.CardSubType) > -1 && p.Volume == MP_Volumn select p).ToList();

            int[,] vs = new int[Station_Count, 2];
            foreach (var item in R_Data)
            {
                var mDate = item.EntryTime;
                if (mDate < OperationDate) continue;

                // 0x00 =>
                // 0x00(0) = 兩個都不加,
                // 0x01(1) = 只加出站,
                // 0x10(2) = 只加進站,
                // 0x11(3) = 兩個都加
                int Item_Row;
                if (IO_Add == MobileEnum.AddEntry || IO_Add == MobileEnum.AddBoth)
                {
                    Item_Row = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == item.EntryStation);
                    vs[Item_Row, 0]++;
                }
                if (IO_Add == MobileEnum.AddExit || IO_Add == MobileEnum.AddBoth)
                {
                    Item_Row = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == item.ExitStation);
                    vs[Item_Row, 1]++;
                }
            }
            return vs;
        }

        private int[,] Volume_CreatColumn(List<MobilePay> Data, ReportSwitch_1 reportType, string CardKind)
        {
            MyCommond.WriteLog(ThisReceive, $"電支資料分配 - 開始");
            int ArrayRow = 1;
            if (reportType == ReportSwitch_1.Station) { ArrayRow = MainServer.globalCommond.Stock_Lrt_Station.Count; }
            else if (reportType == ReportSwitch_1.Day) { ArrayRow = Convert.ToInt16(Convert.ToDateTime(OperationDate.ToString("yyyy/MM/01")).AddDays(-1).ToString("dd")); }
            int[,] vs = new int[ArrayRow, 2];

            var CardKindList = MainServer.globalCommond.Stock_MobileCardType.FindAll(x => x.CardType == CardKind).ToList();
            MyCommond.WriteLog(ThisReceive, $"支付方式 => {CardKind} 卡種筆數 {CardKindList.Count,4}筆");
            var R_Data = (from p in Data where CardKindList.FindIndex(x => x.CardSubType == p.CardSubType) > -1 && p.Volume != MobilePayVolum.InTime select p).ToList();
            MyCommond.WriteLog(ThisReceive, $"計算運量筆數:{R_Data.Count,8}筆");

            foreach (var item in R_Data)
            {
                if (reportType == ReportSwitch_1.Day)
                {
                    var Item_Day = Convert.ToInt16(item.ExitTime.ToString("dd"));
                    var Item_Hour = Convert.ToInt16(item.ExitTime.ToString("HH"));
                    var Item_Row = Convert.ToInt16((Item_Hour < 5) ? $"{Item_Day - 1}" : $"{Item_Day}");

                    vs[Item_Row, 0]++;
                    vs[Item_Row, 1]++;
                }
                else if (reportType == ReportSwitch_1.Station)
                {
                    if (MethodeName < LrtTxnClientSwitch.Month_AllRideList) { if (item.EntryTime < OperationDate) continue; }

                    // 0x00 =>
                    // 0x00(0) = 兩個都不加,
                    // 0x01(1) = 只加出站,
                    // 0x10(2) = 只加進站,
                    // 0x11(3) = 兩個都加
                    int Item_Row_Entry = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == item.EntryStation); /* */ vs[Item_Row_Entry, 0]++;
                    int Item_Row_Exit_ = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == item.ExitStation); /*  */ vs[Item_Row_Exit_, 1]++;
                }
            }
            MyCommond.WriteLog(ThisReceive, $"電支資料分配 - 結束");
            return vs;
        }

        private int[,] Volume_CreatColumn(List<MobilePay> Data, ReportSwitch_1 reportType, MobileCard MobileType, int EachCount)
        {
            MyCommond.WriteLog(ThisReceive, $"電支資料分配 - 開始");
            int ArrayRow = 1;
            int[,] vs = null;
            if (reportType == ReportSwitch_1.Station) { ArrayRow = MainServer.globalCommond.Stock_Lrt_Station.Count; }
            else if (reportType == ReportSwitch_1.Day) { ArrayRow = Convert.ToInt16(Convert.ToDateTime(OperationDate.ToString("yyyy/MM/01")).AddMonths(1).AddDays(-1).ToString("dd")); }

            List<string> MobileType_T1 = new List<string>();
            if (MobileType == MobileCard.Type)
            {
                MobileType_T1 = (from p in MainServer.globalCommond.Stock_MobileCardType group p by new { p.CardType } into g select g.Key.CardType).ToList();
                vs = new int[ArrayRow, (MobileType_T1.Count * EachCount)];
            }
            else if (MobileType == MobileCard.SubType)
            {
                MobileType_T1 = (from p in MainServer.globalCommond.Stock_MobileCardType group p by new { p.CardSubType } into g select g.Key.CardSubType).ToList();
                vs = new int[ArrayRow, (MobileType_T1.Count * EachCount)];
                MobileType_T1 = new List<string>() { "" };
            }

            int ColumnOffset = 0;
            foreach (var item in MobileType_T1)
            {
                List<string> MobileType_T2 = new List<string>();
                if (item == "")
                {
                    MobileType_T2 = (from p in MainServer.globalCommond.Stock_MobileCardType group p by new { p.CardSubType } into g select g.Key.CardSubType).ToList();
                }
                else
                {
                    MobileType_T2 = (from p in MainServer.globalCommond.Stock_MobileCardType where p.CardType == item select p.CardSubType).ToList();
                }

                foreach (var jtem in MobileType_T2)
                {
                    int RowOffset_1 = 0, RowOffset_2 = 0;

                    var mData = (from p in Data where p.CardSubType == jtem select p).ToList();

                    foreach (var ktem in mData)
                    {


                        if (reportType == ReportSwitch_1.Station)
                        {
                            RowOffset_1 = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == ktem.EntryStation);
                            RowOffset_2 = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == ktem.ExitStation);
                        }
                        else if (reportType == ReportSwitch_1.Day)
                        {
                            var ff = Convert.ToDateTime("05:00:00");
                            var dd = Convert.ToDateTime(ktem.ExitTime.ToString("HH:mm:ss"));
                            RowOffset_1 = (ff < dd) ? Convert.ToInt32(ktem.ExitTime.ToString("dd")) : Convert.ToInt32(ktem.ExitTime.AddDays(-1).ToString("dd"));
                        }



                        vs[RowOffset_1, 0] = 0;

                        if (EachCount > 0) { continue; }
                        if (EachCount > 1) { continue; }
                        if (EachCount > 2) { continue; }
                        if (EachCount > 3) { continue; }
                        if (EachCount > 4) { continue; }
                        if (EachCount > 5) { continue; }

                    }
                }
            }

            MyCommond.WriteLog(ThisReceive, $"電支資料分配 - 結束");
            return null;
        }

        /// <summary>
        /// 僅出站及各票種(信用卡、乘車碼)。
        /// </summary>
        /// <param name="Data">行動支付資料。</param>
        /// <param name="_DayOrStation_">車站、天期。</param>
        /// <param name="VolumeOrAmount">運量、營收。</param>
        /// <returns></returns>
        private int[,] Mobile_2(List<MobilePay> Data, ReportSwitch_1 _DayOrStation_, ReportSwitch_2 VolumeOrAmount)
        {
            MyCommond.WriteLog(ThisReceive, $@"" + 
                $@"{((_DayOrStation_ == ReportSwitch_1.Station) ? "每站" : ((_DayOrStation_ == ReportSwitch_1.Day)    ? "每日" : "A"))}電支" +
                $@"{((VolumeOrAmount == ReportSwitch_2.Volumn)  ? "運量" : ((VolumeOrAmount == ReportSwitch_2.Amount) ? "營收" : "B"))}分析");

            int RowMax = -1;

            if (_DayOrStation_ == ReportSwitch_1.Day) { RowMax = Convert.ToDateTime(MainServer.OperationStartDate.ToString("yyyy/MM/01 05:00:00")).AddMonths(1).AddDays(-1).Day; }
            else if (_DayOrStation_ == ReportSwitch_1.Station) { RowMax = MainServer.globalCommond.Stock_Lrt_Station.Count; }
            else {  }
            if (RowMax < 0) return null;
            int[,] Respone = new int[RowMax, 2];

            foreach (var item1 in Data)
            {
                int R_Offset = -1;
                int C_Offset = -1;
                int AddValue = -1;
                
                try
                {
                         if (_DayOrStation_ == ReportSwitch_1.Day) { var date1 = item1.ExitTime; R_Offset = ((5 < date1.Hour) ? date1.AddDays(0).Day : date1.AddDays(-1).Day); }
                    else if (_DayOrStation_ == ReportSwitch_1.Station) { R_Offset = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == item1.ExitStation); }
                    else { }

                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"R_Offset Error!{ex.Message}");
                } // 找行

                try
                {
                    var SearchCartType = MainServer.globalCommond.Stock_MobileCardType.Find(x => x.CardSubType == item1.CardSubType);
                         if (SearchCartType.CardType == "Credit") { C_Offset = 0; }
                    else if (SearchCartType.CardType == "QrCode") { C_Offset = 1; }
                    else {  }
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"C_Offset Error!{ex.Message}");
                } // 找欄

                try
                {
                         if (VolumeOrAmount == ReportSwitch_2.Volumn) { AddValue = 1; }
                    else if (VolumeOrAmount == ReportSwitch_2.Amount) { AddValue = item1.RealPay; }
                    else { }
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"AddValue Error!{ex.Message}");
                } // 數值

                if (R_Offset < 0 || C_Offset < 0 || AddValue < 0) { MyCommond.WriteLog(ThisReceive, $@"Error Data:{item1.ToString()}"); continue; }

                try
                {
                    Respone[--R_Offset, C_Offset] += AddValue;
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"AddValue Error!{ex.Message}");
                } // 寫入
            }

            return Respone;
        }

        private int[,] PaoCal(List<PaoTicketSaleList> PaoData, List<GroupTicketList> GroupData, ReportSwitch_1 _DayOrStation_, ReportSwitch_2 VolumeOrAmount)
        {
            MyCommond.WriteLog(ThisReceive, $@"" + 
                $@"{((_DayOrStation_ == ReportSwitch_1.Station) ? "每站" : ((_DayOrStation_ == ReportSwitch_1.Day)    ? "每日" : "A"))}人工票" +
                $@"{((VolumeOrAmount == ReportSwitch_2.Volumn)  ? "運量" : ((VolumeOrAmount == ReportSwitch_2.Amount) ? "營收" : "B"))}分析");

            int RowMax = -1;
            int ColumnMax = -1;

            //空間高
            if (_DayOrStation_ == ReportSwitch_1.Day) { RowMax = Convert.ToDateTime(MainServer.OperationStartDate.ToString("yyyy/MM/01 05:00:00")).AddMonths(1).AddDays(-1).Day; }
            else if (_DayOrStation_ == ReportSwitch_1.Station) { RowMax = MainServer.globalCommond.Stock_Lrt_Station.Count; }

            //空間寬
            if (VolumeOrAmount == ReportSwitch_2.Volumn) { ColumnMax = 9; }
            else if (VolumeOrAmount == ReportSwitch_2.Amount) { ColumnMax = 11; }

            if (RowMax < 0) return null;
            int[,] Respone = new int[RowMax, ColumnMax];

            foreach (var item1 in PaoData)
            {
                int R_Offset = -1;
                int[] AddValue = null;

                try
                {
                    if (_DayOrStation_ == ReportSwitch_1.Day)
                    {
                        var date1 = new DateTime(Convert.ToInt16(item1.Operation.Substring(0, 4)), Convert.ToInt16(item1.Operation.Substring(4, 2)), Convert.ToInt16(item1.Operation.Substring(6, 2)));
                        R_Offset = date1.Day - 1;
                    }
                    else if (_DayOrStation_ == ReportSwitch_1.Station)
                    {
                        R_Offset = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == item1.Station);
                    }
                    else { }

                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"R_Offset Error!{ex.Message}");
                } // 找行

                try
                {
                    if (VolumeOrAmount == ReportSwitch_2.Volumn)
                    {
                        AddValue = new int[]
                        {
                            /* 1    單程票    */ Convert.ToInt32(item1.Ticket20) + Convert.ToInt32(item1.Ticket25) + Convert.ToInt32(item1.Ticket30) + Convert.ToInt32(item1.ExchangeTicket20) + Convert.ToInt32(item1.ExchangeTicket25) + Convert.ToInt32(item1.ExchangeTicket30), 
                            /* 2    優待票    */ Convert.ToInt32(item1.Ticket10) + Convert.ToInt32(item1.Ticket12) + Convert.ToInt32(item1.Ticket15), 
                            /* 3   自行車票   */ Convert.ToInt32(item1.TicketBike), 
                            /* 4 一日票(公版) */ (Convert.ToInt32(item1.Ticket_OneDay) + Convert.ToInt32(item1.ExchangeOneDay) + Convert.ToInt32(((item1.SplitTicker != "") ? item1.SplitTicker : "0")) + Convert.ToInt32(((item1.KKDay != "") ? item1.KKDay : "0")) + Convert.ToInt32(((item1.Ntpc != "") ? item1.Ntpc : "0"))) * 6, 
                            /* 5 一日票(電票) */ 0, 
                            /* 6 一日票(活動) */ (Convert.ToInt32(item1.Lanwair1) * 6), 
                            /* 7   一日聯票   */ (Convert.ToInt32(item1.Lanwair2) * 6), 
                            /* 8 團體票(張數) */ 0, 
                            /* 9 團體票(人數) */ 0, 
                        };
                    }
                    else if (VolumeOrAmount == ReportSwitch_2.Amount)
                    {
                        AddValue = new int[]
                        {
                            /* 01    單程票    */ ((Convert.ToInt32(item1.Ticket20) * 20) + (Convert.ToInt32(item1.Ticket25) * 25) + (Convert.ToInt32(item1.Ticket30) * 30)), 
                            /* 02    優待票    */ ((Convert.ToInt32(item1.Ticket10) * 10) + (Convert.ToInt32(item1.Ticket12) * 12) + (Convert.ToInt32(item1.Ticket15) * 15)), 
                            /* 03   自行車票   */ (Convert.ToInt32(item1.TicketBike) * 50),  
                            /* 04 一日票(公版) */ (Convert.ToInt32(item1.Ticket_OneDay) * 50), 
                            /* 05 一日票(電票) */ 0, 
                            /* 06 一日票(活動) */ (Convert.ToInt32(item1.Lanwair1) * 50), 
                            /* 07   一日聯票   */ (Convert.ToInt32(item1.Lanwair2) * 50), 
                            /* 08    團體票    */ Convert.ToInt32(item1.GroupAmt), 
                            /* 09   紀念套票   */ Convert.ToInt32(item1.OpenTrafficSetTicket), 
                            /* 10  補票(50倍)  */ Convert.ToInt32(item1.RepairTicket), 
                            /* 11     社福1    */ 0, 
                            /* 12     社福2    */ 0, 
                            /* 13     代墊     */ Convert.ToInt32(item1.ConcessionRefund), 
                        };
                    }
                    else { }
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"AddValue Error!{ex.Message}");
                } // 數值

                if (R_Offset < 0 && AddValue == null) { MyCommond.WriteLog(ThisReceive, $@"Error Data:{item1.ToString()}"); continue; }

                for (int i = 0; i < Respone.GetLength(1); i++) { Respone[R_Offset, i] += AddValue[i]; }
            }
            if (VolumeOrAmount == ReportSwitch_2.Volumn)
            {
                foreach (var item1 in GroupData)
                {
                    int R_Offset = -1;
                    int AddValue = -1;

                    try
                    {
                        if (_DayOrStation_ == ReportSwitch_1.Day)
                        {
                            var date1 = new DateTime(Convert.ToInt16(item1.UseDate.Substring(0, 3)), Convert.ToInt16(item1.UseDate.Substring(3, 2)), Convert.ToInt16(item1.UseDate.Substring(5, 2)));
                            R_Offset = date1.Day - 1;
                        }
                        else if (_DayOrStation_ == ReportSwitch_1.Station)
                        {
                            R_Offset = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == item1.Exit);
                        }
                        else { }

                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $"R_Offset Error!{ex.Message}");
                    } // 找行

                    try
                    {
                             if (VolumeOrAmount == ReportSwitch_2.Volumn) { AddValue = Convert.ToInt32(item1.People); }
                        else if (VolumeOrAmount == ReportSwitch_2.Amount) { }
                        else { }
                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $"AddValue Error!{ex.Message}");
                    } // 數值

                    Respone[R_Offset, 7]++;
                    Respone[R_Offset, 8] += AddValue;

                }
            }
            
            return Respone;
        }


        /// <summary>
        /// 僅出站及各票種(信用卡、乘車碼)。
        /// </summary>
        /// <param name="Data">行動支付資料。</param>
        /// <param name="_DayOrStation_">車站、天期。</param>
        /// <param name="VolumeOrAmount">運量、營收。</param>
        /// <returns></returns>
        private int[,] Mobile_IO(List<MobilePay> Data, ReportSwitch_1 _DayOrStation_, ReportSwitch_2 VolumeOrAmount)
        {
            MyCommond.WriteLog(ThisReceive, $@"" + 
                $@"{((_DayOrStation_ == ReportSwitch_1.Station) ? "每站" : ((_DayOrStation_ == ReportSwitch_1.Day)    ? "每日" : "A"))}電支" +
                $@"{((VolumeOrAmount == ReportSwitch_2.Volumn)  ? "運量" : ((VolumeOrAmount == ReportSwitch_2.Amount) ? "營收" : "B"))}分析");

            int RowMax = -1;

            if (_DayOrStation_ == ReportSwitch_1.Day) { RowMax = Convert.ToDateTime(MainServer.OperationStartDate.ToString("yyyy/MM/01 05:00:00")).AddMonths(1).AddDays(-1).Day; }
            else if (_DayOrStation_ == ReportSwitch_1.Station) { RowMax = MainServer.globalCommond.Stock_Lrt_Station.Count; }
            else {  }
            if (RowMax < 0) return null;
            int[,] Respone = new int[RowMax, 2];

            foreach (var item1 in Data)
            {
                int R_Offset_1 = -1;
                int C_Offset_1 = -1;
                int AddValue_1 = -1;

                int R_Offset_2 = -1;
                int C_Offset_2 = -1;
                int AddValue_2 = -1;

                try
                {
                         if (_DayOrStation_ == ReportSwitch_1.Day)
                    {
                        var date1 = item1.EntryTime;
                        var date2 = item1.ExitTime;
                        R_Offset_1 = ((5 <= date1.Hour) ? date1.Day - 1 : date1.Day - 2);
                        R_Offset_2 = ((5 <= date2.Hour) ? date2.Day - 1 : date2.Day - 2);
                    }
                    else if (_DayOrStation_ == ReportSwitch_1.Station)
                    {
                        R_Offset_1 = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == item1.EntryStation);
                        R_Offset_2 = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == item1.ExitStation);
                    }
                    else { }

                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"R_Offset Error!{ex.Message}");
                } // 找行

                try
                {
                    var SearchCartType = MainServer.globalCommond.Stock_MobileCardType.Find(x => x.CardSubType == item1.CardSubType);
                         if (SearchCartType.CardType == "Credit")
                    {
                        C_Offset_1 = 0;
                        C_Offset_2 = 1;
                    }
                    else if (SearchCartType.CardType == "QrCode")
                    {
                        C_Offset_1 = 2;
                        C_Offset_2 = 3;
                    }
                    else {  }
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"C_Offset Error!{ex.Message}");
                } // 找欄

                try
                {
                         if (VolumeOrAmount == ReportSwitch_2.Volumn)
                    {
                        AddValue_1 = 1;
                        AddValue_2 = 1;
                    }
                    else if (VolumeOrAmount == ReportSwitch_2.Amount)
                    {
                        AddValue_1 = 0;
                        AddValue_2 = item1.RealPay;
                    }
                    else { }
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"AddValue Error!{ex.Message}");
                } // 數值

                if (R_Offset_1 < 0 && C_Offset_1 < 0 && AddValue_1 < 0 && R_Offset_2 < 0 && C_Offset_2 < 0 && AddValue_2 < 0) { MyCommond.WriteLog(ThisReceive, $@"Error Data:{item1.ToString()}"); continue; }

                Respone[R_Offset_1, C_Offset_1] += AddValue_1;
                Respone[R_Offset_2, C_Offset_2] += AddValue_2;
            }
            return null;
        }



        private int[,] Amount_Mobile_2_Line(List<MobilePay> Data, ReportSwitch_1 reportType)
        {
            int[,] vs = null;
            var _cardType = (from p in MainServer.globalCommond.Stock_MobileCardType group p by p.CardType into g select g.Key).ToList();

            int newColumn = _cardType.Count;
            if (reportType == ReportSwitch_1.Day)
            {
                var sss = Convert.ToInt32(Convert.ToDateTime(OperationDate.ToString("yyyy/MM/01")).AddMonths(1).AddDays(-1).ToString("dd"));
                vs = new int[sss, newColumn];
            }
            else if (reportType == ReportSwitch_1.Station)
            {
                vs = new int[MainServer.globalCommond.Stock_Lrt_Station.Count, newColumn];
            }
            if (vs == null) return null;
            MyCommond.WriteLog(ThisReceive, $"Row:{vs.GetLength(0),4}, Column:{vs.GetLength(1),4}");
            MyCommond.WriteLog(ThisReceive, $"電支交易方式共{_cardType.Count,4}種");
            List<int[,]> respone = new List<int[,]>();
            int Column = 0;
            foreach (var item in _cardType)
            {
                MyCommond.WriteLog(ThisReceive, $"電支交易方式:{item}");
                int ColumnOffset = Column;
                var bbb = (from p in MainServer.globalCommond.Stock_MobileCardType where p.CardType == item group p by p.CardSubType into g select g.Key).ToList();
                MyCommond.WriteLog(ThisReceive, $"此交易方式共{bbb.Count,4}家票證公司");
                var Data_2 = (from p in Data where bbb.Contains(p.CardSubType) select p).ToList();
                MyCommond.WriteLog(ThisReceive, $"共篩選出{Data_2.Count,8}筆交易");
                foreach (var jtem in Data_2)
                {
                    try
                    {
                        MyCommond.WriteLog(ThisReceive, $"取得:{jtem}");
                        int RowOffset = 0;
                        if (reportType == ReportSwitch_1.Day)
                        {
                            if (true)
                            {
                                if (jtem.ExitTime.Hour < 5)
                                {
                                    RowOffset = jtem.ExitTime.AddDays(-1).Day - 1;
                                }
                                else
                                {
                                    RowOffset = jtem.ExitTime.Day - 1;
                                }
                            }
                            else
                            {
                                var ff = Convert.ToDateTime("05:00:00");
                                var dd = Convert.ToDateTime(jtem.ExitTime.ToString("HH:mm:ss"));
                                var mYear = jtem.ExitTime.Year;
                                var mMonth = jtem.ExitTime.Month;
                                var mDay = jtem.ExitTime.Day;
                                var mHour = jtem.ExitTime.Hour;
                                var mMinute = jtem.ExitTime.Minute;
                                var mSecond = jtem.ExitTime.Second;
                                //RowOffset = (ff < dd) ? Convert.ToInt32(jtem.ExitTime.ToString("dd")) : Convert.ToInt32(jtem.ExitTime.AddDays(-1).ToString("dd"));
                                RowOffset = (5 < mHour) ? (mDay - 1) : (mDay - 2);
                            }
                            MyCommond.WriteLog(ThisReceive, $"天期:{RowOffset}");

                        }
                        else if (reportType == ReportSwitch_1.Station)
                        {
                            RowOffset = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == jtem.ExitStation);
                            MyCommond.WriteLog(ThisReceive, $"車站:{RowOffset}");
                        }

                        if (RowOffset == -1) continue;
                        MyCommond.WriteLog(ThisReceive, $"vs[{RowOffset,4}, {ColumnOffset,4}] = {vs[RowOffset, ColumnOffset]} + {jtem.RealPay}");
                        vs[RowOffset, ColumnOffset] += Convert.ToInt32(jtem.RealPay);
                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $"Error:{ex.Message}");
                    }
                }
                Column++;
            }
            return vs;
        }

        private int[,] Volume_Mobile_2_2_Line(List<MobilePay> Data, ReportSwitch_1 reportType)
        {
            MyCommond.WriteLog(ThisReceive, $"開始相關團票資料擷取");
            var _cardType = (from p in MainServer.globalCommond.Stock_MobileCardType group p by p.CardType into g select g.Key).ToList();
            int[,] vs = new int[MainServer.globalCommond.Stock_Lrt_Station.Count, _cardType.Count * 2];
            MyCommond.WriteLog(ThisReceive, $"Row:{vs.GetLength(0),4}, Column:{vs.GetLength(1),4}");
            MyCommond.WriteLog(ThisReceive, $"電支交易方式共{_cardType.Count,4}種");
            List<int[,]> respone = new List<int[,]>();
            int Column = 0;
            foreach (var item in _cardType)
            {
                MyCommond.WriteLog(ThisReceive, $"電支交易方式:{item}");
                int ColumnOffset = Column * 2;
                var bbb = (from p in MainServer.globalCommond.Stock_MobileCardType where p.CardType == item group p by p.CardSubType into g select g.Key).ToList();
                MyCommond.WriteLog(ThisReceive, $"此交易方式共{bbb.Count,4}家票證公司");
                var Data_2 = (from p in Data where bbb.Contains(p.CardSubType) select p).ToList();
                MyCommond.WriteLog(ThisReceive, $"共篩選出{Data_2.Count,8}筆交易");
                foreach (var jtem in Data_2)
                {
                    try
                    {
                        MyCommond.WriteLog(ThisReceive, $"取得:{jtem}");
                        int IndexEntry = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == jtem.EntryStation);
                        int IndexExit_ = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == jtem.ExitStation);

                        MyCommond.WriteLog(ThisReceive, $"vs[{IndexEntry,4}, {ColumnOffset + 0,4}] = {vs[IndexEntry, ColumnOffset + 0]} + 1");
                        MyCommond.WriteLog(ThisReceive, $"vs[{IndexExit_,4}, {ColumnOffset + 1,4}] = {vs[IndexExit_, ColumnOffset + 1]} + 1");
                        vs[IndexEntry, ColumnOffset + 0]++;
                        vs[IndexExit_, ColumnOffset + 1]++;
                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $"Error:{ex.Message}");
                    }
                }
                Column++;
            }
            return vs;
        }

        private int[,] Volume_Mobile_5_Line(List<MobilePay> Data, ReportSwitch_1 reportType)
        {
            int[,] vs = null;
            var _cardSubType = (from p in MainServer.globalCommond.Stock_MobileCardType group p by p.CardSubType into g select g.Key).ToList();
            int newColumn = _cardSubType.Count;
            if (reportType == ReportSwitch_1.Day)
            {
                var sss = Convert.ToInt32(Convert.ToDateTime(OperationDate.ToString("yyyy/MM/01")).AddMonths(1).AddDays(-1).ToString("dd"));
                vs = new int[sss, newColumn];
            }
            else if (reportType == ReportSwitch_1.Station)
            {
                vs = new int[MainServer.globalCommond.Stock_Lrt_Station.Count, newColumn];
            }
            if (vs == null) return null;
            MyCommond.WriteLog(ThisReceive, $"Row:{vs.GetLength(0),4}, Column:{vs.GetLength(1),4}");
            MyCommond.WriteLog(ThisReceive, $"此交易方式共{_cardSubType.Count,4}家票證公司");
            List<int[,]> respone = new List<int[,]>();
            int Column = 0;
            foreach (var item in _cardSubType)
            {
                MyCommond.WriteLog(ThisReceive, $"家票證公司:{item}");
                int ColumnOffset = Column;
                var Data_2 = (from p in Data where p.CardSubType == item select p).ToList();
                MyCommond.WriteLog(ThisReceive, $"共篩選出{Data_2.Count,8}筆交易");
                foreach (var jtem in Data_2)
                {
                    try
                    {
                        MyCommond.WriteLog(ThisReceive, $"取得:{jtem}");
                        int RowOffset = 0;
                        if (reportType == ReportSwitch_1.Day)
                        {
                            if (true)
                            {
                                if (jtem.ExitTime.Hour < 5)
                                {
                                    RowOffset = jtem.ExitTime.AddDays(-1).Day - 1;
                                }
                                else
                                {
                                    RowOffset = jtem.ExitTime.Day - 1;
                                }
                            }
                            else
                            {
                                var ff = Convert.ToDateTime("05:00:00");
                                var dd = Convert.ToDateTime(jtem.ExitTime.ToString("HH:mm:ss"));
                                var mYear = jtem.ExitTime.Year;
                                var mMonth = jtem.ExitTime.Month;
                                var mDay = jtem.ExitTime.Day;
                                var mHour = jtem.ExitTime.Hour;
                                var mMinute = jtem.ExitTime.Minute;
                                var mSecond = jtem.ExitTime.Second;
                                //RowOffset = (ff < dd) ? Convert.ToInt32(jtem.ExitTime.ToString("dd")) : Convert.ToInt32(jtem.ExitTime.AddDays(-1).ToString("dd"));
                                RowOffset = (5 < mHour) ? (mDay - 1) : (mDay - 2);
                            }
                            MyCommond.WriteLog(ThisReceive, $"天期:{RowOffset}");

                        }
                        else if (reportType == ReportSwitch_1.Station)
                        {
                            RowOffset = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == jtem.ExitStation);
                            MyCommond.WriteLog(ThisReceive, $"車站:{RowOffset}");
                        }

                        if (RowOffset == -1) continue;
                        MyCommond.WriteLog(ThisReceive, $"vs[{RowOffset,4}, {ColumnOffset,4}] = {vs[RowOffset, ColumnOffset]} + 1");
                        vs[RowOffset, ColumnOffset]++;
                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $"Error:{ex.Message}");
                    }
                }
                Column++;
            }
            return vs;
        }

        private int[,] Amount_10_Line(List<MobilePay> Data, ReportSwitch_1 reportType)
        {
            int[,] vs = null;
            var _cardSubType = (from p in MainServer.globalCommond.Stock_MobileCardType group p by p.CardSubType into g select g.Key).ToList();
            int newColumn = _cardSubType.Count * 2;
            if (reportType == ReportSwitch_1.Day)
            {
                var sss = Convert.ToInt32(Convert.ToDateTime(OperationDate.ToString("yyyy/MM/01")).AddMonths(1).AddDays(-1).ToString("dd"));
                vs = new int[sss, newColumn];
            }
            else if (reportType == ReportSwitch_1.Station)
            {
                vs = new int[MainServer.globalCommond.Stock_Lrt_Station.Count, newColumn];
            }
            if (vs == null) return null;
            MyCommond.WriteLog(ThisReceive, $"Row:{vs.GetLength(0),4}, Column:{vs.GetLength(1),4}");
            MyCommond.WriteLog(ThisReceive, $"此交易方式共{_cardSubType.Count,4}家票證公司");
            List<int[,]> respone = new List<int[,]>();
            int Column = 0;
            foreach (var item in _cardSubType)
            {
                MyCommond.WriteLog(ThisReceive, $"家票證公司:{item}");
                int ColumnOffset = Column * 2;
                var Data_2 = (from p in Data where p.CardSubType == item select p).ToList();
                MyCommond.WriteLog(ThisReceive, $"共篩選出{Data_2.Count,8}筆交易");
                foreach (var jtem in Data_2)
                {
                    try
                    {
                        MyCommond.WriteLog(ThisReceive, $"取得:{jtem}");
                        int RowOffset = 0;
                        if (reportType == ReportSwitch_1.Day)
                        {
                            if (true)
                            {
                                if (jtem.ExitTime.Hour < 5)
                                {
                                    RowOffset = jtem.ExitTime.AddDays(-1).Day - 1;
                                }
                                else
                                {
                                    RowOffset = jtem.ExitTime.Day - 1;
                                }
                            }
                            else
                            {
                                var ff = Convert.ToDateTime("05:00:00");
                                var dd = Convert.ToDateTime(jtem.ExitTime.ToString("HH:mm:ss"));
                                var mYear = jtem.ExitTime.Year;
                                var mMonth = jtem.ExitTime.Month;
                                var mDay = jtem.ExitTime.Day;
                                var mHour = jtem.ExitTime.Hour;
                                var mMinute = jtem.ExitTime.Minute;
                                var mSecond = jtem.ExitTime.Second;
                                //RowOffset = (ff < dd) ? Convert.ToInt32(jtem.ExitTime.ToString("dd")) : Convert.ToInt32(jtem.ExitTime.AddDays(-1).ToString("dd"));
                                RowOffset = (5 < mHour) ? (mDay - 1) : (mDay - 2);
                            }
                            MyCommond.WriteLog(ThisReceive, $"天期:{RowOffset}");
                        }
                        else if (reportType == ReportSwitch_1.Station)
                        {
                            RowOffset = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == jtem.ExitStation);
                            MyCommond.WriteLog(ThisReceive, $"車站:{RowOffset}");
                        }

                        if (RowOffset == -1) continue;
                        MyCommond.WriteLog(ThisReceive, $"vs[{RowOffset,4}, {ColumnOffset + 0,4}] = {vs[RowOffset, ColumnOffset + 0]} + {jtem.ShouldPay}");
                        MyCommond.WriteLog(ThisReceive, $"vs[{RowOffset,4}, {ColumnOffset + 1,4}] = {vs[RowOffset, ColumnOffset + 1]} + {jtem.Fine}");
                        vs[RowOffset, ColumnOffset + 0] += Convert.ToInt32(jtem.ShouldPay);
                        vs[RowOffset, ColumnOffset + 1] += Convert.ToInt32(jtem.Fine);
                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $"Error:{ex.Message}");
                    }
                }
                Column++;
            }
            return vs;
        }

        private int[,] Volume_AddList(List<int[,]> vs)
        {
            MyCommond.WriteLog(ThisReceive, $"將List中的二陣列合併");
            int[,] Respon = null;
            foreach (var item in vs)
            {
                if (Respon == null)
                {
                    MyCommond.WriteLog(ThisReceive, $"迴圈開始，初始化變數Respon");
                    Respon = new int[item.GetLength(0), item.GetLength(1)];
                }
                for (int i = 0; i < item.GetLength(0); i++)
                {
                    for (int j = 0; j < item.GetLength(1); j++)
                    {
                        MyCommond.WriteLog(ThisReceive, $"Respon[{i,4}, {j,4}] => {Respon[i, j],8} + {item[i, j],8} = { Respon[i, j] += item[i, j],8}");
                        //Respon[i, j] += item[i, j];
                    }
                }
            }
            return Respon;
        }

        private int[,] Volume_ListJoin(List<int[,]> vs)
        {
            int[,] Respon = null;
            int SwitchCount = 0;
            foreach (var item in vs)
            {
                if (Respon == null)
                {
                    MyCommond.WriteLog(ThisReceive, $"初始化變數");
                    Respon = new int[item.GetLength(0), item.GetLength(1) * 2];
                }
                for (int i = 0; i < item.GetLength(0); i++)
                {
                    for (int j = 0; j < item.GetLength(1); j++)
                    {
                        MyCommond.WriteLog(ThisReceive, $"Respon[{i,4}, {j + SwitchCount * item.GetLength(1),4}] => {Respon[i, j + SwitchCount * item.GetLength(1)],8} + {item[i, j],8} = {Respon[i, j + SwitchCount * item.GetLength(1)] += item[i, j],8}");
                        //Respon[i, j + SwitchCount * item.GetLength(1)] += item[i, j];
                    }
                }
                SwitchCount++;
            }
            return Respon;
        }

        private int[,] Amount_CreatColumn(string CardKind)
        {
            var Station_Count = MainServer.globalCommond.Stock_Lrt_Station.Count;
            var CardKindList = MainServer.globalCommond.Stock_MobileCardType.FindAll(x => x.CardType == CardKind).ToList();
            MyCommond.WriteLog(ThisReceive, $"支付方式 => {CardKind} 卡種筆數 {CardKindList.Count,4}筆");
            var R_Data = (from p in Day_MobilePayData where CardKindList.FindIndex(x => x.CardSubType == p.CardSubType) > -1 select p).ToList();
            MyCommond.WriteLog(ThisReceive, $"計算運量筆數:{R_Data.Count,8}筆");

            int[,] Respon = new int[Station_Count, 1];
            foreach (var item in R_Data)
            {
                var mDate = item.EntryTime;
                if (mDate < OperationDate) continue;

                // 0x00 =>
                // 0x00(0) = 兩個都不加,
                // 0x01(1) = 只加出站,
                // 0x10(2) = 只加進站,
                // 0x11(3) = 兩個都加
                int Item_Row = MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.StationName == item.ExitStation);
                Respon[Item_Row, 0] += item.RealPay;
            }
            return Respon;
        }

        #endregion

        #region 輸出

        private string[] inputZ(string[] data)
        {
            int RC = data.Length;
            string[] result = new string[RC];
            for (int i = 0; i < RC; i++)
            {
                result[i] = ((data[i] != "") ? data[i]:"0");
            }
            return result;
        }

        private string[,] inputZ(string[,] data)
        {
            int RC = data.GetLength(0);
            int CC = data.GetLength(1);
            string[,] result = new string[RC, CC];
            for(int i = 0; i < RC; i++)
            {
                for (int j = 0; j < CC; j++)
                {
                    result[i, j] = ((data[i, j] != "") ? data[i, j] : "0");
                }
            }
            return result;
        }

        /** private void ConsoleExportArray(int[,] result, string Title = "")
        {
            if (result == null)
            {
                MyCommond.WriteLog(ThisReceive, $"---Null---");
                return;
            }
            if (Title != "")
            {
                MyCommond.WriteLog(ThisReceive, $"---FullAdd({Title})---");
            }
            else
            {
                MyCommond.WriteLog(ThisReceive, $"---FullAdd---");
            }

            for (int i = 0; i < result.GetLength(0); i++)
            {
                string LineString = "";
                for (int j = 0; j < result.GetLength(1); j++)
                {
                    LineString += $"{((result[i, j] != null) ? result[i, j] : 0),4}, ";
                }
                MyCommond.WriteLog(ThisReceive, $"vs[{i,4}] = {LineString}");
            }
        }
        */
        /** private void ConsoleExportArray(string[,] result, string Title = "")
        {
            if (result == null)
            {
                MyCommond.WriteLog(ThisReceive, $"---Null---");
                return;
            }
            if (Title != "")
            {
                MyCommond.WriteLog(ThisReceive, $"---FullAdd({Title})---");
            }
            else
            {
                MyCommond.WriteLog(ThisReceive, $"---FullAdd---");
            }

            for (int i = 0; i < result.GetLength(0); i++)
            {
                string LineString = "";
                for (int j = 0; j < result.GetLength(1); j++)
                {
                    LineString += $"{((result[i, j] != null) ? result[i, j] : ""),8}, ";
                }
                MyCommond.WriteLog(ThisReceive, $"vs[{i,4}] = {LineString}");
            }
        }
        */
        /** private int[,] TResultToArray<TResult>(List<TResult> InputData)
        {
            MyCommond.WriteLog(ThisReceive, $"將資料表轉換成多行陣列 - 開始");

            ValueCount += 1;

            var Count_R = InputData.Count;
            var Count_C = InputData[0].ToString().Split(',').Length;
            int[,] OutputData = new int[Count_R, Count_C];
            int Run_R = 0;
            foreach (var item in InputData)
            {
                var LineData = item.ToString().Split(',');
                for (int Run_C = 0; Run_C < LineData.Length; Run_C++)
                {
                    OutputData[Run_R, Run_C] = Convert.ToInt32(LineData[Run_C]);
                }
            }

            ValueCount += 1;

            MyCommond.WriteLog(ThisReceive, $"將資料表轉換成多行陣列 - 結束");
            return OutputData;
        }
        */
        /** private int[,] ArrayToExport(int[,] Data)
        {
            MyCommond.WriteLog(ThisReceive, "提取內容資料 - 開始");

            ValueCount += 1;

            int[,] Array = null;
            int RowMax = Data.GetLength(0); ;
            int ColumnMax;
            int ColumnOffset = 0;

            int CountIfDate = 0;
            int CountIfStation = 0;

            if (Data[0, 0] == OperationDate.Year)
            {
                MyCommond.WriteLog(ThisReceive, $"年分 位移 + 1");
                CountIfDate++;
                if (Data[0, 1] == OperationDate.Month)
                {
                    MyCommond.WriteLog(ThisReceive, $"月分 位移 + 1");
                    CountIfDate++;
                    if (Data[0, 2] == OperationDate.Day)
                    {
                        MyCommond.WriteLog(ThisReceive, $"日期 位移 + 1");
                        CountIfDate++;
                        if (Data[0, 3] == OperationDate.Hour)
                        {
                            MyCommond.WriteLog(ThisReceive, $"小時 位移 + 1");
                            CountIfDate++;
                        }
                    }
                }
            }
            else
            {
                for (int a = 0; a < Data.GetLength(0); a++)
                {
                    bool SearchData_Station = (MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.CodeNumber == Convert.ToString(Data[a, 0])) >= 0) ? true : false;
                    if (SearchData_Station) CountIfStation++;
                }
            }

            ValueCount += 1;

            if (CountIfDate > 0)
            {
                MyCommond.WriteLog(ThisReceive, $"此為日期 位移量:{CountIfDate}");
                ColumnOffset = CountIfDate;
            }
            else if (CountIfStation == RowMax)
            {
                MyCommond.WriteLog(ThisReceive, $"此為車站 位移量:{CountIfStation}");
                ColumnOffset = 1;
            }
            else return null;

            ValueCount += 1;

            ColumnMax = Data.GetLength(1) - ColumnOffset;
            MyCommond.WriteLog(ThisReceive, $"原變數行數:{Data.GetLength(0),4}, 欄數:{Data.GetLength(1),4}");
            MyCommond.WriteLog(ThisReceive, $"新變數行數:{RowMax,4}, 欄數:{ColumnMax,4}");
            Array = new int[RowMax, ColumnMax];

            for (int i = 0; i < RowMax; i++)
            {
                for (int j = 0; j < ColumnMax; j++)
                {
                    Array[i, j] = Data[i, j + ColumnOffset];
                }
            }

            ValueCount += 1;

            MyCommond.WriteLog(ThisReceive, "提取內容資料 - 完成");
            return Array;
        }
        */
        /** public bool ExportExcel_Old(ExportConfig ec, string[] mTitle, string[,] mData)
        {
            string[] TextWay = new string[] { "csv", "txt", "config" };
            string[] ExcelWay = new string[] { "xls", "xlsx" };

            string filePath = ec.Path1;
            string fileName1 = ec.DTime.ToString("yyyyMMdd");
            string fileName2 = ec.FirstName;
            string fileName3 = ec.FileName;
            string fileName4 = ec.SubName;

            string FullFileName = ((fileName1.Length > 0) ? $"{fileName1}_" : "") +
                ((fileName2.Length > 0) ? $"{fileName2}_" : "") +
                ((fileName3.Length > 0) ? $"{fileName3}_" : "") +
                ((fileName4.Length > 0) ? $"{fileName4}_" : "") +
                ec.OpLine;

            string ThisExportFullFileName = ExportFileName;
            string FullPath = filePath + ThisExportFullFileName;


            MyCommond.WriteLog(ThisReceive, $"開始輸出文字檔");
            for (int i = 0; i < mData.GetLength(0); i++)
            {
                string ExportString = "";
                string ExportDataRow = "";
                for (int j = 0; j < mData.GetLength(1); j++)
                {
                    if (i == 0) ExportString += mTitle[j] + ",";
                    ExportDataRow += mData[i, j] + ",";
                }

                if (!File.Exists(FullPath)) File.WriteAllText(FullPath, ExportString + Environment.NewLine, Encoding.UTF8);
                else File.AppendAllText(FullPath, ExportString + Environment.NewLine, Encoding.UTF8);

                if (i == 0)
                {
                    ExportString = ExportDataRow;
                    if (!File.Exists(FullPath)) File.WriteAllText(FullPath, ExportString + Environment.NewLine, Encoding.UTF8);
                    else File.AppendAllText(FullPath, ExportString + Environment.NewLine, Encoding.UTF8);
                }

            }

            return false;
        }
        */
        /**private int[,] TResultToArray<TResult>(TResult InputData)
        {
            MyCommond.WriteLog(ThisReceive, $"將資料表轉換成單行陣列 - 開始");

            ValueCount += 1;

            var Count_C = InputData.ToString().Split(',').Length;
            int[,] OutputData = new int[1, Count_C];
            var LineData = InputData.ToString().Split(',');
            for (int Run_C = 0; Run_C < LineData.Length; Run_C++)
            {
                OutputData[0, Run_C] = Convert.ToInt32(LineData[Run_C]);
            }

            ValueCount += 1;

            MyCommond.WriteLog(ThisReceive, $"將資料表轉換成單行陣列 - 結束");
            return OutputData;
        }
        */

        private int[,] ArrayToExport(string[,] Data)
        {
            MyCommond.WriteLog(ThisReceive, "提取內容資料 - 開始");

            ValueCount += 1;

            MyCommond.WriteLog(ThisReceive, $"原始資料長度:{Data.GetLength(0),8}, 寬度:{Data.GetLength(1),8}");
            int[,] Array = null;
            int RowMax = Data.GetLength(0); ;
            int ColumnMax;
            int ColumnOffset = 0;

            int CountIfDate = 0;
            int CountIfStation = 0;

            if (Convert.ToInt32(Data[0, 0]) == OperationDate.Year)
            {
                MyCommond.WriteLog(ThisReceive, $"年分 位移 + 1");
                CountIfDate++;
                if (Convert.ToInt32(Data[0, 1]) == OperationDate.Month)
                {
                    MyCommond.WriteLog(ThisReceive, $"月分 位移 + 1");
                    CountIfDate++;
                    if (Convert.ToInt32(Data[0, 2]) == OperationDate.Day)
                    {
                        MyCommond.WriteLog(ThisReceive, $"日期 位移 + 1");
                        CountIfDate++;
                        if (Convert.ToInt32(Data[0, 3]) == OperationDate.Hour)
                        {
                            MyCommond.WriteLog(ThisReceive, $"小時 位移 + 1");
                            CountIfDate++;
                        }
                    }
                }
            }
            else
            {
                for (int a = 0; a < Data.GetLength(0); a++)
                {
                    bool SearchData_Station = (MainServer.globalCommond.Stock_Lrt_Station.FindIndex(x => x.CodeNumber == Convert.ToString(Data[a, 0])) >= 0) ? true : false;
                    if (SearchData_Station) CountIfStation++;
                }
            }

            ValueCount += 1;

            if (CountIfDate > 0)
            {
                MyCommond.WriteLog(ThisReceive, $"此為日期 位移量:{CountIfDate}");
                ColumnOffset = CountIfDate;
            }
            else if (CountIfStation == RowMax)
            {
                MyCommond.WriteLog(ThisReceive, $"此為車站 位移量:{CountIfStation}");
                ColumnOffset = 1;
            }
            else
            {
                MyCommond.WriteLog(ThisReceive, $"直接將字串輸出成整數");
                ColumnOffset = 0;
            }

            ValueCount += 1;

            ColumnMax = Data.GetLength(1) - ColumnOffset;
            MyCommond.WriteLog(ThisReceive, $"原變數行數:{Data.GetLength(0),4}, 欄數:{Data.GetLength(1),4}");
            MyCommond.WriteLog(ThisReceive, $"新變數行數:{RowMax,4}, 欄數:{ColumnMax,4}");
            Array = new int[RowMax, ColumnMax];

            MyCommond.ConsoleExportArray(ThisReceive, MethodeName.ToString(), Data);

            for (int i = 0; i < RowMax; i++)
            {
                for (int j = 0; j < ColumnMax; j++)
                {
                    int sss = Convert.ToInt32((Data[i, j + ColumnOffset] == "") ? "0" : Data[i, j + ColumnOffset]);
                    MyCommond.WriteLog(ThisReceive, $"Array[{i,4}, {j,4}] = {sss}");
                    Array[i, j] = sss;
                }
            }

            ValueCount += 1;

            MyCommond.WriteLog(ThisReceive, "提取內容資料 - 完成");
            return Array;
        }

        public bool ExportExcle_V1(string template, int[,] Data_Afc, int[,] Data_MobilePay, int[,] Data_OwnTicket)
        {

            #region 讀取輸出位置資料

            MyCommond.WriteLog(ThisReceive, $"取得輸出資料變數");
            string mYear = "", mMonth = "", mDay = "";

            bool ExportCheck = false;
            bool Customized = (MethodeName == LrtTxnClientSwitch.Day_Volume || MethodeName == LrtTxnClientSwitch.Day_Amount/* || MethodeName == LrtTxnClientSwitch.Day_EachStation_EachTime*/);

            var S_R_F = MainServer.globalCommond.Stock_ReportFile.FindAll(x => x.OperationLine == MainServer.globalCommond.OperationCode).ToList();
            var S_R_F_Index = MainServer.globalCommond.Stock_ReportFile.FindIndex(x => x.OperationLine == MainServer.globalCommond.OperationCode);
            MyCommond.WriteLog(ThisReceive, $"找到 {((S_R_F_Index > -1) ? $"{S_R_F.Count}":"0"),4} 筆資料");
            if (S_R_F_Index == -1) { return false; }


            var MethodeName_zhTW = MethodeName.ToDescription();
            var MethodeName_String = MethodeName.ToString();

            int ReportDateStartRow /*----*/ = Convert.ToInt16((/*----------*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).ReportDateRow == "-") /*-----*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).ReportDateRow)); /*-----*/ MyCommond.WriteLog(ThisReceive, "產製日期起始行：" + ReportDateStartRow);
            int ReportDateStartColumn /*-*/ = MyCommond.ExcelColumnNum((/*-*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).ReportDateColumn == "-") /*--*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).ReportDateColumn)); /*--*/ MyCommond.WriteLog(ThisReceive, "產製日期起始欄：" + ReportDateStartColumn);
            int StartRow /*--------------*/ = Convert.ToInt16((/*----------*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).StartRow == "-") /*----------*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).StartRow)); /*----------*/ MyCommond.WriteLog(ThisReceive, "輸出資料起始行：" + StartRow);
            int StartColumn /*-----------*/ = MyCommond.ExcelColumnNum((/*-*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).StartColumn == "-") /*-------*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).StartColumn)); /*-------*/ MyCommond.WriteLog(ThisReceive, "輸出資料起始欄：" + StartColumn);
            int EndRow /*----------------*/ = Convert.ToInt16((/*----------*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).EndRow == "-") /*------------*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).EndRow)); /*------------*/ MyCommond.WriteLog(ThisReceive, "輸出資料結束行：" + EndRow);
            int EndColumn /*-------------*/ = MyCommond.ExcelColumnNum((/*-*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).EndColumn == "-") /*---------*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).EndColumn)); /*---------*/ MyCommond.WriteLog(ThisReceive, "輸出資料結束欄：" + EndColumn);
            int thisBreakOffRow /*-------*/ = Convert.ToInt16((/*----------*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffRow == "-") /*-------*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffRow)); /*-------*/ MyCommond.WriteLog(ThisReceive, "資料中斷行    ：" + thisBreakOffRow);
            int thisBreakOffRowNum /*----*/ = Convert.ToInt16((/*----------*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffRowNum == "-") /*----*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffRowNum)); /*----*/ MyCommond.WriteLog(ThisReceive, "資料中斷行數  ：" + thisBreakOffRowNum);
            int thisBreakOffColumn /*----*/ = MyCommond.ExcelColumnNum((/*-*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffColumn == "-") /*----*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffColumn)); /*----*/ MyCommond.WriteLog(ThisReceive, "資料中斷欄    ：" + thisBreakOffColumn);
            int thisBreakOffColumnNum /*-*/ = Convert.ToInt16((/*----------*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffColumnNum == "-") /*-*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffColumnNum)); /*-*/ MyCommond.WriteLog(ThisReceive, "資料中斷欄數  ：" + thisBreakOffColumnNum);
            int MobilePayColumn /*-------*/ = MyCommond.ExcelColumnNum((/*-*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).MobilePayColumn == "-") /*---*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).MobilePayColumn)); /*---*/ MyCommond.WriteLog(ThisReceive, "電子支付起始欄：" + MobilePayColumn);
            int OwnTiketColumn /*--------*/ = MyCommond.ExcelColumnNum((/*-*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).OwnTiketColumn == "-") /*----*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).OwnTiketColumn)); /*----*/ MyCommond.WriteLog(ThisReceive, "自有票起始欄  ：" + OwnTiketColumn);

            mYear = Convert.ToString(Convert.ToInt32(this.OperationDate.ToString("yyyy")) - 1911);
            mMonth = Convert.ToString(this.OperationDate.ToString("MM"));
            mDay = (MethodeName < LrtTxnClientSwitch.Month_AllRideList) ? Convert.ToString(this.OperationDate.ToString("dd")) : "";

            #endregion

            if (Customized) { goto CustomizedSub; }
            else { goto NormalSub; }

        CustomizedSub:
            MyCommond.WriteLog(ThisReceive, $"客製化運量、營收輸出 - 開始");

            #region 運量、營收

            ValueCount += 1;

            while (MainServer.VA_Lock) { Thread.Sleep(1000); }

            MainServer.VA_Lock = true;

            Int32 Workbook_Table;
            Int32 YearAgo = 0, MonthAgo = 0, LastMonthDayCount = 0, ThisDay = 0;
            Int32 RowOffset = 0;
            Int32 SumValue;
            Excel.Range myRange;
            string strRange = "";
            DateTime StrDate;
            bool IsFirstYear, IsNextYear, IsNextMonth;

            if (MethodeName == LrtTxnClientSwitch.Day_Volume) { Workbook_Table = 1; }
            else if (MethodeName == LrtTxnClientSwitch.Day_Amount) { Workbook_Table = 2; }
            else { return false; }

            ValueCount += 1;

            #region Offset

            YearAgo = OperationDate.AddDays(-0).Year - MainServer.globalCommond.Operation_First_Day.Year;
            MyCommond.WriteLog(ThisReceive, $"變數YearAgo狀態          :{Convert.ToString(YearAgo)}");
            LastMonthDayCount = Convert.ToInt32(DateTime.DaysInMonth(OperationDate.AddDays(-1).Year, OperationDate.AddDays(-1).Month));  //計算跨月前的月份天數
            MyCommond.WriteLog(ThisReceive, $"變數LastMonthDayCount狀態:{Convert.ToString(LastMonthDayCount)}");
            ThisDay = Convert.ToInt16(OperationDate.AddDays(-0).ToString("dd"));
            MyCommond.WriteLog(ThisReceive, $"變數ThisDay狀態          :{Convert.ToString(ThisDay)}");

            IsFirstYear = OperationDate.AddDays(-1).Year == MainServer.globalCommond.Operation_First_Day.Year;
            MyCommond.WriteLog(ThisReceive, $"變數IsFirstYear狀態      :{Convert.ToString(IsFirstYear)}");
            IsNextYear = OperationDate.AddDays(-1).Year != OperationDate.AddDays(-0).Year;
            MyCommond.WriteLog(ThisReceive, $"變數IsNextYear狀態       :{Convert.ToString(IsNextYear)}");
            IsNextMonth = OperationDate.AddDays(-1).Month != OperationDate.AddDays(-0).Month;
            MyCommond.WriteLog(ThisReceive, $"變數IsNextMonth狀態      :{Convert.ToString(IsNextMonth)}");

            MonthAgo = OperationDate.AddDays(-1).Month;
            MyCommond.WriteLog(ThisReceive, $"變數MonthAgo起始狀態     :{Convert.ToString(MonthAgo)}");
            if (IsFirstYear)
            {
                MonthAgo = MonthAgo - MainServer.globalCommond.Operation_First_Day.Month + 1;
                MyCommond.WriteLog(ThisReceive, $"變數MonthAgo狀態         :{Convert.ToString(MonthAgo)}");
            }
            MonthAgo--;
            ValueCount += 1;

            #endregion

            string FileName_DayRideAmt = @"{0}_營運資料統計表(總運量、總營收)_{1}.xlsx";         //1120101_營運資料統計表(總運量、總營收).xlsx

            string BookFileName, BookYear, BookMonth, BookDay;
            string NewBookFileName, NewBookYear, NewBookMonth, NewBookDay;

            //設定開啟檔名
            BookYear = Convert.ToString(Convert.ToInt16(OperationDate.AddDays(-1).ToString("yyyy")) - 1911);
            BookMonth = Convert.ToString(OperationDate.AddDays(-1).ToString("MM"));
            BookDay = Convert.ToString(OperationDate.AddDays(-1).ToString("dd"));
            BookFileName = String.Format(FileName_DayRideAmt, BookYear + BookMonth + BookDay, MainServer.globalCommond.OperationCode_String);

            NewBookYear = Convert.ToString(Convert.ToInt16(OperationDate.AddDays(0).ToString("yyyy")) - 1911);
            NewBookMonth = Convert.ToString(OperationDate.AddDays(0).ToString("MM"));
            NewBookDay = Convert.ToString(OperationDate.AddDays(0).ToString("dd"));
            NewBookFileName = String.Format(FileName_DayRideAmt, NewBookYear + NewBookMonth + NewBookDay, MainServer.globalCommond.OperationCode_String);

            MyCommond.WriteLog(ThisReceive, $"變數BookFileName狀態     :{Convert.ToString(BookFileName)}");
            MyCommond.WriteLog(ThisReceive, $"變數NewBookFileName狀態  :{Convert.ToString(NewBookFileName)}");

            ValueCount += 1;

            Excel.Application xlApp = new Excel.Application();
            try
            {
                //開啟檔案
                MyCommond.WriteLog(ThisReceive, $"開啟Excel Application");
                xlApp.Visible = MainServer.globalCommond.Excel_Applaction_Show;
                string filenamepath = MyCommond.Path_Program + MyCommond.Path_Report;
                bool fileExist = File.Exists(filenamepath + NewBookFileName);
                MyCommond.WriteLog(ThisReceive, $"新檔案是否存在:{fileExist}");
                Excel.Workbook OperationWB;
                if (fileExist) { OperationWB = xlApp.Workbooks.Open(Filename: filenamepath + NewBookFileName, ReadOnly: false); }
                else { OperationWB = xlApp.Workbooks.Open(Filename: filenamepath + BookFileName, ReadOnly: false); }

                ValueCount += 1;

                MyCommond.WriteLog(ThisReceive, $"開啟Excel第 {Convert.ToString(Workbook_Table)} 籤頁");
                Excel.Worksheet ws2 = (Excel.Worksheet)OperationWB.Worksheets[Workbook_Table];

                ValueCount += 1;

                #region 月合併
                if (IsNextMonth)
                {
                    if (IsFirstYear) { RowOffset = MonthAgo - MainServer.globalCommond.Operation_First_Day.Month + 2; }
                    else if (YearAgo > 3)  { RowOffset = 3 + MonthAgo; }
                    else { RowOffset = YearAgo + MonthAgo; }

                    MyCommond.WriteLog(ThisReceive, $"變數RowOffset狀態        :{Convert.ToString(RowOffset)}");
                    int tgR = StartRow + RowOffset;
                    int tgC = StartColumn;
                    string msgStr;

                    ValueCount += 1;

                    MyCommond.WriteLog(ThisReceive, $"月合併開始");
                    while (tgC <= EndColumn)
                    {
                        MyCommond.WriteLog(ThisReceive, $"開始處理Column第 {tgC} 欄的資料");
                        if (tgC == StartColumn)
                        {
                            MyCommond.WriteLog(ThisReceive, $"在Row第 {Convert.ToString(tgR)} 行前插入一行");
                            ws2.Rows[tgR].insert();
                            //取得要合併的欄位
                            strRange = MyCommond.ExcelColumnStr(tgC) + (tgR) + ":" + MyCommond.ExcelColumnStr(tgC + 1) + (tgR);
                            MyCommond.WriteLog(ThisReceive, $"指定合併範圍 {strRange}");
                            //合併欄位
                            MyCommond.WriteLog(ThisReceive, $"將該範圍合併、文字置中、將格式定義為YYYY.MM");
                            ws2.Range[strRange].Merge();
                            //將合併欄位的文字置中
                            ws2.Cells[tgR, tgC].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            //修改格式
                            ws2.Cells[tgR, tgC].NumberFormatLocal = "yyyy.mm";
                            //取得日期
                            StrDate = Convert.ToDateTime(ws2.Cells[tgR + 1, tgC].text);
                            MyCommond.WriteLog(ThisReceive, $"取得第 {Convert.ToString(tgR + 1)} 行日期為 {StrDate.ToString()}");
                            //輸入營運日前的日期
                            MyCommond.WriteLog(ThisReceive, $"將日期寫入 {MyCommond.ExcelColumnStr(tgC)}{Convert.ToString(tgR)}");
                            ws2.Cells[tgR, tgC] = StrDate.ToString("yyyy/MM/01");
                        }   //寫入該月份
                        else if (tgC == StartColumn + 1)
                        {
                            MyCommond.WriteLog(ThisReceive, $"此欄不做處理");
                        }   //星期
                        else if (tgC == StartColumn + 2)
                        {
                            if (MethodeName == LrtTxnClientSwitch.Day_Volume)
                            {
                                ws2.Cells[tgR, tgC] = $"=SUM(E{tgR}:{MyCommond.ExcelColumnStr(EndColumn)}{tgR})-S{tgR}";
                                msgStr = string.Format($"=SUM(E{tgR}:{MyCommond.ExcelColumnStr(EndColumn)}{tgR})-S{tgR}");
                            }
                            else if (MethodeName == LrtTxnClientSwitch.Day_Amount)
                            {
                                ws2.Cells[tgR, tgC] = $"=SUM(E{tgR}:{MyCommond.ExcelColumnStr(EndColumn)}{tgR})";
                                msgStr = string.Format($"=SUM(E{tgR}:{MyCommond.ExcelColumnStr(EndColumn)}{tgR})");
                            }
                            else
                            {
                                msgStr = "啥?";
                            }

                            MyCommond.WriteLog(ThisReceive, $"{msgStr}");
                        }   //總計
                        else if (tgC >= (StartColumn + 3) && tgC <= EndColumn)
                        {
                            //設定範圍
                            strRange = MyCommond.ExcelColumnStr(tgC) + (tgR + 1) + ":" + MyCommond.ExcelColumnStr(tgC) + (tgR + LastMonthDayCount);
                            //取得要計算的範圍
                            myRange = ws2.Range[strRange];
                            //將範圍內的值做加總
                            SumValue = Convert.ToInt32(xlApp.WorksheetFunction.Sum(myRange));
                            msgStr = string.Format(@"將範圍 {0} 加總和為:{1} 寫入{2}{3} ,並清除範圍資料", strRange, SumValue, MyCommond.ExcelColumnStr(tgC), Convert.ToString(tgR));
                            MyCommond.WriteLog(ThisReceive, $"{msgStr}");
                            //將加總的值輸入
                            ws2.Cells[tgR, tgC] = SumValue.ToString();
                            //清除資料
                            myRange.ClearContents();
                        }   //所有票分別加總
                        tgC++;
                    }
                    //設定本月第一天日期
                    MyCommond.WriteLog(ThisReceive, $"在 {MyCommond.ExcelColumnStr(StartColumn)}:{Convert.ToString(tgR + 1)} 寫入總和 {Convert.ToString(OperationDate.AddDays(-0).ToString("yyyy/MM/dd"))}");
                    ws2.Cells[tgR + 1, StartColumn] = OperationDate.AddDays(-0).ToString("yyyy/MM/dd");
                    MyCommond.WriteLog(ThisReceive, $"月合併結束");
                }
                #endregion

                ValueCount += 1;

                #region 年合併
                if (IsNextYear)
                {
                    string msgStr;
                    if (YearAgo > 3) { RowOffset = 3; } else { RowOffset = YearAgo; }
                    MyCommond.WriteLog(ThisReceive, $"年合併開始");
                    int tgR = StartRow + RowOffset - 1;
                    int tgC = StartColumn;
                    msgStr = "將{0}:{1}寫到{0}:{2}";

                    ValueCount += 1;

                    //填寫或搬移年總運量(保留近3年)
                    while (tgC <= EndColumn)
                    {
                        if (tgC == StartColumn)
                        {
                            if (YearAgo <= 3)
                            {
                                MyCommond.WriteLog(ThisReceive, $"在第 {Convert.ToString(tgR)}前插入一行");
                                ws2.Rows[tgR].insert();
                                //取得要合併的欄位
                                strRange = MyCommond.ExcelColumnStr(tgC) + (tgR) + ":" + MyCommond.ExcelColumnStr(tgC + 1) + (tgR);
                                //合併欄位
                                ws2.Range[strRange].Merge();
                                //將合併欄位的文字置中
                                ws2.Cells[tgR, tgC].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            }
                            else
                            {
                                ws2.Cells[tgR - 2, tgC] = ws2.Cells[tgR - 1, tgC];
                                ws2.Cells[tgR - 1, tgC] = ws2.Cells[tgR, tgC];
                            }
                            //取得日期
                            StrDate = Convert.ToDateTime(ws2.Cells[tgR + 1, tgC].text);
                            //將日期重新編排輸入
                            ws2.Cells[tgR, tgC].NumberFormatLocal = "yyyy年";
                            ws2.Cells[tgR, tgC] = StrDate.ToString("yyyy/01/01");
                            //msgStr = "取得 " + MyCommond.ExcelColumnStr(tgC) + ":" + Convert.ToString(tgR + 1) + " 資料 " + StrDate + " 寫入 " + MyCommond.ExcelColumnStr(tgC) + Convert.ToString(tgR) + "並將設改為YYYY.MM";
                            msgStr = string.Format("取得{0}:{1}資料{2}寫入{3}{4}，並將資料型態設為yyyy.MM", MyCommond.ExcelColumnStr(tgC), Convert.ToString(tgR + 1), StrDate, MyCommond.ExcelColumnStr(tgC), Convert.ToString(tgR));
                            MyCommond.WriteLog(ThisReceive, $"{msgStr}");
                        }   //寫入該月份
                        else if (tgC == StartColumn + 1)
                        {
                            MyCommond.WriteLog(ThisReceive, $"此欄不做處理");
                        }   //星期
                        else if (tgC == StartColumn + 2)
                        {
                            if (MethodeName == LrtTxnClientSwitch.Day_Volume)
                            {
                                ws2.Cells[tgR, tgC] = $"=SUM(E{tgR}:{MyCommond.ExcelColumnStr(EndColumn)}{tgR})-S{tgR}";
                                msgStr = string.Format($"=SUM(E{tgR}:{MyCommond.ExcelColumnStr(EndColumn)}{tgR})-S{tgR}");
                            }
                            else if (MethodeName == LrtTxnClientSwitch.Day_Amount)
                            {
                                ws2.Cells[tgR, tgC] = $"=SUM(E{tgR}:{MyCommond.ExcelColumnStr(EndColumn)}{tgR})";
                                msgStr = string.Format($"=SUM(E{tgR}:{MyCommond.ExcelColumnStr(EndColumn)}{tgR})");
                            }
                            else
                            {
                                msgStr = "啥?";
                            }
                            MyCommond.WriteLog(ThisReceive, $"{msgStr}");
                        }   //總計
                        else if (tgC > (StartColumn + 2) && tgC <= EndColumn)
                        {
                            if (YearAgo > 3)
                            {
                                ws2.Cells[tgR - 2, tgC] = ws2.Cells[tgR - 1, tgC];
                                ws2.Cells[tgR - 1, tgC] = ws2.Cells[tgR, tgC];
                            }
                            //設定範圍
                            strRange = MyCommond.ExcelColumnStr(tgC) + (tgR + 1) + ":" + MyCommond.ExcelColumnStr(tgC) + (tgR + 1 + MonthAgo);
                            //取得要計算的範圍
                            myRange = ws2.Range[strRange];
                            //將範圍內的值做加總
                            SumValue = Convert.ToInt32(xlApp.WorksheetFunction.Sum(myRange));
                            msgStr = string.Format("將範圍 {0} 加總，和為:{1} 寫入{2}{3}", strRange, SumValue, MyCommond.ExcelColumnStr(tgC), Convert.ToString(tgR));
                            MyCommond.WriteLog(ThisReceive, $"msgStr");
                            //將加總的值輸入
                            ws2.Cells[tgR, tgC] = SumValue.ToString();
                        }   //所有票分別加總
                        tgC++;
                    }

                    ValueCount += 1;

                    MyCommond.WriteLog(ThisReceive, $"開始移除每月運量");
                    //移除每月總運量
                    Int32 WrithTime = 0;
                    strRange = (tgR + 1) + ":" + ((tgR + 1) + MonthAgo);
                    ws2.Rows[strRange].delete();
                }
                #endregion

                //加入後修改
                MyCommond.WriteLog(ThisReceive, $"開始寫入每日資料");
                Int32 WriteRow, WriteColumn;

                IsFirstYear = OperationDate.AddDays(-0).Year == MainServer.globalCommond.Operation_First_Day.Year;
                MyCommond.WriteLog(ThisReceive, $"變更IsFirstYear狀態      :{Convert.ToString(IsFirstYear)}");

                MonthAgo = Convert.ToDateTime(OperationDate.ToString("yyyy/MM/01")).AddMonths(-1).Month;
                if (MonthAgo == 12) { MonthAgo = 0; }

                if (IsFirstYear) MonthAgo = MonthAgo - MainServer.globalCommond.Operation_First_Day.Month;
                MyCommond.WriteLog(ThisReceive, $"變數MonthAgo狀態:{Convert.ToString(MonthAgo)}");

                ValueCount += 1;

                if (IsFirstYear) WriteRow = StartRow + MonthAgo + ThisDay;
                else if (YearAgo > 3) WriteRow = StartRow + 3 + (MonthAgo - 1) + ThisDay;
                else WriteRow = StartRow + YearAgo + (MonthAgo - 1) + ThisDay;
                MyCommond.WriteLog(ThisReceive, $"寫入起始行:{WriteRow}");

                ValueCount += 1;

                WriteColumn = StartColumn + 3;
                MyCommond.WriteLog(ThisReceive, $"寫入起始欄:{WriteColumn}");

                string msgStr_Y = "第 {0} 行第 {1} 欄寫入第 {2} 比資料：{3}";
                for (var i = 0; i < Data_Afc.GetLength(0); i++)
                {
                    for (var ColumnOffset2 = 0; ColumnOffset2 < Data_Afc.GetLength(1); ColumnOffset2++)
                    {
                        MyCommond.WriteLog(ThisReceive, $"{string.Format(msgStr_Y, WriteRow, WriteColumn + ColumnOffset2, i * ColumnOffset2, Data_Afc[i, ColumnOffset2])}");
                        ws2.Cells[WriteRow, WriteColumn + ColumnOffset2] = Data_Afc[i, ColumnOffset2];
                    }
                }

                ValueCount += 1;

                //設定儲存檔名
                if (fileExist)
                {
                    MyCommond.WriteLog(ThisReceive, $"儲存檔案");
                    OperationWB.Save();
                }
                else
                {
                    MyCommond.WriteLog(ThisReceive, $"另存成:{NewBookFileName}");
                    ws2.SaveAs(filenamepath + NewBookFileName);
                }

                ValueCount += 1;

                MyCommond.WriteLog(ThisReceive, $"關閉檔案");
                OperationWB.Close();
                MyCommond.WriteLog(ThisReceive, $"釋放Excel物件");
                Marshal.ReleaseComObject(ws2);
                Marshal.ReleaseComObject(OperationWB);
                ExportCheck = true;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{MyCommond.ExportErrorMessageToLog(ex, MethodeName.ToDescription()).Message}");
                ExportCheck = false;
            }
            finally
            {
                MyCommond.WriteLog(ThisReceive, $"關閉Excel Application");
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                MainServer.VA_Lock = false;
            }

            ValueCount += 1;

            #endregion

            MyCommond.WriteLog(ThisReceive, $"客製化運量、營收輸出 - 結束");
            return ExportCheck;

        NormalSub:
            MyCommond.WriteLog(ThisReceive, $"其他通用輸出 - 開始");

            #region 普通

            ValueCount += 1;

            bool ckeck = MyCommond.CheckFile(MyCommond.Path_Program + MyCommond.Path_Report + ExportFileName);
            XLWorkbook wb = new XLWorkbook();
            if (ckeck)
            {
                MyCommond.WriteLog(ThisReceive, "以既有的檔案複寫");
                wb = new XLWorkbook(MyCommond.Path_Program + MyCommond.Path_Report + ExportFileName);
            }
            else
            {
                if (!(MyCommond.CheckFile(template)))
                {
                    MyCommond.WriteLog(ThisReceive, "無樣版檔，認後再重新執行");

                    Thread thread = new Thread(() => MessageBox.Show($"無樣版檔\n{template}\n認後再重新執行", "錯誤"));
                    thread.IsBackground = true;
                    thread.Start();
                    return ExportCheck;
                }

                MyCommond.WriteLog(ThisReceive, "以樣板檔案寫入");
                wb = new XLWorkbook(template);
            }

            ValueCount += 1;

            try
            {
                var ws = wb.Worksheets.First();
                int RowBreak, ColumnBreak;

                MyCommond.WriteLog(ThisReceive, "寫入產製日期");
                if (ReportDateStartRow > 0 && ReportDateStartColumn > 0)
                {
                    ws.Cell(ReportDateStartRow, ReportDateStartColumn).Value = $"{OperationDate}";
                    if (!(MethodeName == LrtTxnClientSwitch.Month_Volume || MethodeName == LrtTxnClientSwitch.Month_Amount))
                    {
                        ws.Cell(ReportDateStartRow + 1, ReportDateStartColumn).Value = DateTime.Now.ToString("yyyy/MM/dd");
                    }
                }
                
                ValueCount += 1;

                #region 一般資料寫入

                try
                {
                    MyCommond.WriteLog(ThisReceive, $"Data_Afc -> Row:{Data_Afc.GetLength(0),4}, Column:{Data_Afc.GetLength(1),4}");
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"Data_Afc -> Error\n{ex.Message}");
                }
                finally
                {

                }

                if (Data_Afc != null && Data_Afc.GetLength(0) > 0 && Data_Afc.GetLength(1) > 0)
                {
                    if (thisBreakOffRow != 0)
                    {
                        RowBreak = thisBreakOffRow - StartRow + 1;
                        EndRow = EndRow - StartRow - thisBreakOffRowNum;
                    }
                    else
                    {
                        RowBreak = Data_Afc.GetLength(0);
                    }
                    if (thisBreakOffColumn != 0)
                    {
                        ColumnBreak = thisBreakOffColumn - StartColumn + 1;
                        EndColumn = EndColumn - StartColumn - thisBreakOffColumnNum + 1;
                    }
                    else
                    {
                        ColumnBreak = Data_Afc.GetLength(1);
                    }

                    MyCommond.WriteLog(ThisReceive, $"---RowBreak:{RowBreak,4}, EndRow:{EndRow,4}, ColumnBreak:{ColumnBreak,4}, EndColumn:{EndColumn,4}---");

                    ValueCount += 1;

                    MyCommond.WriteLog(ThisReceive, "開始資料寫入");
                    for (int RowOffset_1 = 0; RowOffset_1 < RowBreak; RowOffset_1++)
                    {
                        try
                        {
                            for (int ColumnOffset = 0; ColumnOffset < ColumnBreak; ColumnOffset++)
                            {
                                MyCommond.WriteLog(ThisReceive, $"" +
                                    $"sheet.Cell({(StartRow + RowOffset_1),4}, {(StartColumn + ColumnOffset),4}).Value = " +
                                    $"{Data_Afc[RowOffset_1, ColumnOffset]}");
                                ws.Cell(StartRow + RowOffset_1, StartColumn + ColumnOffset).Value = Data_Afc[RowOffset_1, ColumnOffset];
                            }
                            if (thisBreakOffColumn != 0)
                            {
                                for (int ColumnOffset = ColumnBreak; ColumnOffset < EndColumn; ColumnOffset++)
                                {
                                    MyCommond.WriteLog(ThisReceive, $"" +
                                        $"sheet.Cell({(StartRow + RowOffset_1),4}, {(StartColumn + thisBreakOffColumnNum + ColumnOffset),4}).Value = " +
                                        $"{Data_Afc[RowOffset_1, ColumnOffset]}");
                                    ws.Cell(StartRow + RowOffset_1, StartColumn + thisBreakOffColumnNum + ColumnOffset).Value = Data_Afc[RowOffset_1, ColumnOffset];
                                }
                            }  //斷點欄後寫數值不對
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"ERROR! \n{MyCommond.ExportErrorMessageToLog(ex, MethodeName.ToDescription()).Message}");
                        }
                    }

                    MyCommond.WriteLog(ThisReceive, "行中斷資料寫入");
                    if (thisBreakOffRow != 0)
                    {
                        for (int RowOffset_2 = RowBreak; RowOffset_2 < EndRow; RowOffset_2++)
                        {
                            try
                            {
                                for (int ColumnOffset = 0; ColumnOffset < ColumnBreak; ColumnOffset++)
                                {
                                    MyCommond.WriteLog(ThisReceive, $"" +
                                        $"sheet.Cell({(StartRow + thisBreakOffRowNum + RowOffset_2),4}, {(StartColumn + ColumnOffset),4}).Value = " +
                                        $"{Data_Afc[RowOffset_2, ColumnOffset]}");
                                    ws.Cell(StartRow + thisBreakOffRowNum + RowOffset_2, StartColumn + ColumnOffset).Value = Data_Afc[RowOffset_2, ColumnOffset];
                                }
                                if (thisBreakOffColumn != 0)
                                {
                                    for (int ColumnOffset = ColumnBreak; ColumnOffset < EndColumn; ColumnOffset++)
                                    {
                                        MyCommond.WriteLog(ThisReceive, $"" +
                                            $"sheet.Cell({(StartRow + thisBreakOffRowNum + RowOffset_2),4}, {(StartColumn + thisBreakOffColumnNum + ColumnOffset),4}).Value = " +
                                            $"{Data_Afc[RowOffset_2, ColumnOffset]}");
                                        ws.Cell(StartRow + thisBreakOffRowNum + RowOffset_2, StartColumn + thisBreakOffColumnNum + ColumnOffset).Value = Data_Afc[RowOffset_2, ColumnOffset];
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                MyCommond.WriteLog(ThisReceive, $"ERROR! \n{MyCommond.ExportErrorMessageToLog(ex, MethodeName.ToDescription()).Message}");
                            }
                        }
                    }
                }

                ValueCount += 1;

                #endregion

                #region 行動支付寫入

                try
                {
                    MyCommond.WriteLog(ThisReceive, $"Data_MobilePay -> Row:{Data_MobilePay.GetLength(0),4}, Column:{Data_MobilePay.GetLength(1),4}");
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"Data_MobilePay -> Error\n{ex.Message}");
                }
                finally
                {

                }

                if (Data_MobilePay != null && Data_MobilePay.GetLength(0) > 0 && Data_MobilePay.GetLength(1) > 0)
                {
                    MyCommond.WriteLog(ThisReceive, "行動支付寫入");
                    for (int RowOffset_3 = 0; RowOffset_3 < Data_MobilePay.GetLength(0); RowOffset_3++)
                    {
                        for (int ColumnOffset_3 = 0; ColumnOffset_3 < Data_MobilePay.GetLength(1); ColumnOffset_3++)
                        {
                            int RowSet = StartRow + RowOffset_3;
                            int ColumnSet = MobilePayColumn + ColumnOffset_3;
                            //int ExcelCellValue = Convert.ToInt32(Convert.ToString(ws.Cell(RowSet, ColumnSet).Value) ?? "0");
                            int ExcelCellValue = ((ws.Cell(RowSet, ColumnSet).Value == "") ? 0 : Convert.ToInt32(ws.Cell(RowSet, ColumnSet).Value));
                            MyCommond.WriteLog(ThisReceive, $"" +
                                $"sheet.Cell({RowSet}, {ColumnSet}).Value = " +
                                $"{ExcelCellValue} + {Data_MobilePay[RowOffset_3, ColumnOffset_3]} = " +
                                $"{ExcelCellValue + Data_MobilePay[RowOffset_3, ColumnOffset_3]}");
                            ws.Cell(RowSet, ColumnSet).Value = Convert.ToInt64(ExcelCellValue) + Data_MobilePay[RowOffset_3, ColumnOffset_3];
                        }
                    }
                }

                ValueCount += 1;

                #endregion

                #region 人工販售寫入

                try
                {
                    MyCommond.WriteLog(ThisReceive, $"Data_OwnTicket -> Row:{Data_OwnTicket.GetLength(0),4}, Column:{Data_OwnTicket.GetLength(1),4}");
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"Data_OwnTicket -> Error\n{ex.Message}");
                }
                finally
                {

                }

                if (Data_OwnTicket != null && Data_OwnTicket.GetLength(0) > 0 && Data_OwnTicket.GetLength(1) > 0)
                {
                    MyCommond.WriteLog(ThisReceive, $"自有票寫入");
                    for (int RowOffset_3 = 0; RowOffset_3 < Data_OwnTicket.GetLength(0); RowOffset_3++)
                    {
                        for (int ColumnOffset_3 = 0; ColumnOffset_3 < Data_OwnTicket.GetLength(1); ColumnOffset_3++)
                        {
                            if (Data_OwnTicket[RowOffset_3, ColumnOffset_3] == null) continue;
                            int RowSet = StartRow + RowOffset_3;
                            int ColumnSet = OwnTiketColumn + ColumnOffset_3;
                            MyCommond.WriteLog(ThisReceive, $"" +
                                $"sheet.Cell({RowSet}, {ColumnSet}).Value = " +
                                $"{((ws.Cell(RowSet, ColumnSet).Value != "") ? Convert.ToInt64(ws.Cell(RowSet, ColumnSet).Value) : 0)} + {Data_OwnTicket[RowOffset_3, ColumnOffset_3]} = " +
                                $"{((ws.Cell(RowSet, ColumnSet).Value != "") ? Convert.ToInt64(ws.Cell(RowSet, ColumnSet).Value) : 0) + Data_OwnTicket[RowOffset_3, ColumnOffset_3]}");
                            ws.Cell(RowSet, ColumnSet).Value = ((ws.Cell(RowSet, ColumnSet).Value != "") ? Convert.ToInt64(ws.Cell(RowSet, ColumnSet).Value) : 0) + Data_OwnTicket[RowOffset_3, ColumnOffset_3];
                        }
                    }
                }

                ValueCount += 1;

                #endregion

                MyCommond.WriteLog(ThisReceive, $"寫入結束，將檔案存成{ExportFileName}");
                ExportCheck = true;
                ValueCount += 1;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR! \n{MyCommond.ExportErrorMessageToLog(ex, MethodeName.ToDescription()).Message}");
            }
            finally
            {
                wb.SaveAs(ExportPath + ExportFileName);
                wb = null;
                MyCommond.WriteLog(ThisReceive, $"結束Excel輸出");
            }

            #endregion

            MyCommond.WriteLog(ThisReceive, $"其他通用輸出 - 結束");

            return ExportCheck;
        }

        public bool ExportExcle_V2(string template, int[,] Data_Afc, int[,] Data_MobilePay, int[,] Data_OwnTicket)
        {
            
            #region 讀取輸出位置資料

            MyCommond.WriteLog(ThisReceive, $"取得輸出資料變數");
            string mYear = "", mMonth = "", mDay = "";

            bool ExportCheck = false;
            bool Customized = (MethodeName == LrtTxnClientSwitch.Day_Volume || MethodeName == LrtTxnClientSwitch.Day_Amount/* || MethodeName == LrtTxnClientSwitch.Day_EachStation_EachTime*/);

            var S_R_F = MainServer.globalCommond.Stock_ReportFile.FindAll(x => x.OperationLine == MainServer.globalCommond.OperationCode).ToList();
            var S_R_F_Index = MainServer.globalCommond.Stock_ReportFile.FindIndex(x => x.OperationLine == MainServer.globalCommond.OperationCode);
            MyCommond.WriteLog(ThisReceive, $"找到 {((S_R_F_Index > -1) ? $"{S_R_F.Count}" : "0"),4} 筆資料");
            if (S_R_F_Index == -1) { return false; }


            var MethodeName_zhTW = MethodeName.ToDescription();
            var MethodeName_String = MethodeName.ToString();

            int ReportDateStartRow /*----*/ = Convert.ToInt16((/*----------*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).ReportDateRow == "-") /*-----*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).ReportDateRow)); /*-----*/ MyCommond.WriteLog(ThisReceive, "產製日期起始行：" + ReportDateStartRow);
            int ReportDateStartColumn /*-*/ = MyCommond.ExcelColumnNum((/*-*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).ReportDateColumn == "-") /*--*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).ReportDateColumn)); /*--*/ MyCommond.WriteLog(ThisReceive, "產製日期起始欄：" + ReportDateStartColumn);
            int StartRow /*--------------*/ = Convert.ToInt16((/*----------*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).StartRow == "-") /*----------*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).StartRow)); /*----------*/ MyCommond.WriteLog(ThisReceive, "輸出資料起始行：" + StartRow);
            int StartColumn /*-----------*/ = MyCommond.ExcelColumnNum((/*-*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).StartColumn == "-") /*-------*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).StartColumn)); /*-------*/ MyCommond.WriteLog(ThisReceive, "輸出資料起始欄：" + StartColumn);
            int EndRow /*----------------*/ = Convert.ToInt16((/*----------*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).EndRow == "-") /*------------*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).EndRow)); /*------------*/ MyCommond.WriteLog(ThisReceive, "輸出資料結束行：" + EndRow);
            int EndColumn /*-------------*/ = MyCommond.ExcelColumnNum((/*-*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).EndColumn == "-") /*---------*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).EndColumn)); /*---------*/ MyCommond.WriteLog(ThisReceive, "輸出資料結束欄：" + EndColumn);
            int thisBreakOffRow /*-------*/ = Convert.ToInt16((/*----------*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffRow == "-") /*-------*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffRow)); /*-------*/ MyCommond.WriteLog(ThisReceive, "資料中斷行    ：" + thisBreakOffRow);
            int thisBreakOffRowNum /*----*/ = Convert.ToInt16((/*----------*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffRowNum == "-") /*----*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffRowNum)); /*----*/ MyCommond.WriteLog(ThisReceive, "資料中斷行數  ：" + thisBreakOffRowNum);
            int thisBreakOffColumn /*----*/ = MyCommond.ExcelColumnNum((/*-*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffColumn == "-") /*----*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffColumn)); /*----*/ MyCommond.WriteLog(ThisReceive, "資料中斷欄    ：" + thisBreakOffColumn);
            int thisBreakOffColumnNum /*-*/ = Convert.ToInt16((/*----------*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffColumnNum == "-") /*-*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffColumnNum)); /*-*/ MyCommond.WriteLog(ThisReceive, "資料中斷欄數  ：" + thisBreakOffColumnNum);
            int MobilePayColumn /*-------*/ = MyCommond.ExcelColumnNum((/*-*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).MobilePayColumn == "-") /*---*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).MobilePayColumn)); /*---*/ MyCommond.WriteLog(ThisReceive, "電子支付起始欄：" + MobilePayColumn);
            int OwnTiketColumn /*--------*/ = MyCommond.ExcelColumnNum((/*-*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).OwnTiketColumn == "-") /*----*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).OwnTiketColumn)); /*----*/ MyCommond.WriteLog(ThisReceive, "自有票起始欄  ：" + OwnTiketColumn);

            mYear = Convert.ToString(Convert.ToInt32(this.OperationDate.ToString("yyyy")) - 1911);
            mMonth = Convert.ToString(this.OperationDate.ToString("MM"));
            mDay = (MethodeName < LrtTxnClientSwitch.Month_AllRideList) ? Convert.ToString(this.OperationDate.ToString("dd")) : "";

            #endregion 讀取輸出位置資料

            if (Customized) { goto CustomizedSub; }
            else { goto NormalSub; }

        CustomizedSub:
            MyCommond.WriteLog(ThisReceive, $"客製化運量、營收輸出 - 開始");

            #region 運量、營收

            ValueCount += 1;
            var random = new Random();
            Thread.Sleep(1000 * random.Next(11));
            while (MainServer.VA_Lock) { Thread.Sleep(1000 * random.Next(11)); }

            int Reserve_Year = 0;

            MainServer.VA_Lock = true;

            Int32 Workbook_Table;
            Int32 YearAgo = 0, MonthAgo = 0, LastMonthDayCount = 0, ThisDay = 0;
            Int32 RowOffset = 0;
            Int32 SumValue;
            Excel.Range myRange;
            string strRange = "";
            DateTime StrDate;
            bool IsFirstYear, IsNextYear, IsNextMonth;

            if (MethodeName == LrtTxnClientSwitch.Day_Volume) { Workbook_Table = 1; }
            else if (MethodeName == LrtTxnClientSwitch.Day_Amount) { Workbook_Table = 2; }
            else { return false; }

            ValueCount += 1;

            #region Offset

            YearAgo = OperationDate.AddDays(-0).Year - MainServer.globalCommond.Operation_First_Day.Year;
            MyCommond.WriteLog(ThisReceive, $"變數YearAgo狀態          :{Convert.ToString(YearAgo)}");
            LastMonthDayCount = Convert.ToInt32(DateTime.DaysInMonth(OperationDate.AddDays(-1).Year, OperationDate.AddDays(-1).Month));  //計算跨月前的月份天數
            MyCommond.WriteLog(ThisReceive, $"變數LastMonthDayCount狀態:{Convert.ToString(LastMonthDayCount)}");
            ThisDay = Convert.ToInt16(OperationDate.AddDays(-0).ToString("dd"));
            MyCommond.WriteLog(ThisReceive, $"變數ThisDay狀態          :{Convert.ToString(ThisDay)}");

            IsFirstYear = OperationDate.AddDays(-1).Year == MainServer.globalCommond.Operation_First_Day.Year;
            MyCommond.WriteLog(ThisReceive, $"變數IsFirstYear狀態      :{Convert.ToString(IsFirstYear)}");
            IsNextYear = OperationDate.AddDays(-1).Year != OperationDate.AddDays(-0).Year;
            MyCommond.WriteLog(ThisReceive, $"變數IsNextYear狀態       :{Convert.ToString(IsNextYear)}");
            IsNextMonth = OperationDate.AddDays(-1).Month != OperationDate.AddDays(-0).Month;
            MyCommond.WriteLog(ThisReceive, $"變數IsNextMonth狀態      :{Convert.ToString(IsNextMonth)}");

            MonthAgo = OperationDate.AddDays(-1).Month;
            MyCommond.WriteLog(ThisReceive, $"變數MonthAgo起始狀態     :{Convert.ToString(MonthAgo)}");
            if (IsFirstYear)
            {
                MonthAgo = MonthAgo - MainServer.globalCommond.Operation_First_Day.Month + 1;
                MyCommond.WriteLog(ThisReceive, $"變數MonthAgo狀態         :{Convert.ToString(MonthAgo)}");
            }
            //MonthAgo--;
            ValueCount += 1;

            if (YearAgo != 0) { YearAgo = 1; }

            #endregion

            string FileName_DayRideAmt = @"{0}_營運資料統計表(總運量、總營收)_{1}.xlsx";         //1120101_營運資料統計表(總運量、總營收).xlsx

            string BookFileName, BookYear, BookMonth, BookDay;
            string NewBookFileName, NewBookYear, NewBookMonth, NewBookDay;

            var S_R_F_Line = MainServer.globalCommond.Stock_ReportFile.FindAll(x =>
                       x.ReportNameEng == MethodeName_String
                       );

            //設定開啟檔名
            BookYear = Convert.ToString(Convert.ToInt16(OperationDate.AddDays(-1).ToString("yyyy")) - 1911);
            BookMonth = Convert.ToString(OperationDate.AddDays(-1).ToString("MM"));
            BookDay = Convert.ToString(OperationDate.AddDays(-1).ToString("dd"));
            BookFileName = String.Format(FileName_DayRideAmt, BookYear + BookMonth + BookDay, MainServer.globalCommond.OperationCode_String);

            NewBookYear = Convert.ToString(Convert.ToInt16(OperationDate.AddDays(0).ToString("yyyy")) - 1911);
            NewBookMonth = Convert.ToString(OperationDate.AddDays(0).ToString("MM"));
            NewBookDay = Convert.ToString(OperationDate.AddDays(0).ToString("dd"));
            NewBookFileName = String.Format(FileName_DayRideAmt, NewBookYear + NewBookMonth + NewBookDay, MainServer.globalCommond.OperationCode_String);

            EndColumn = S_R_F_Line.Max(x => MyCommond.ExcelColumnNum(x.EndColumn));

            MyCommond.WriteLog(ThisReceive, $"變數BookFileName狀態     :{Convert.ToString(BookFileName)}");
            MyCommond.WriteLog(ThisReceive, $"變數NewBookFileName狀態  :{Convert.ToString(NewBookFileName)}");

            Excel.Application xlApp = new Excel.Application();
            try
            {
                //開啟檔案
                MyCommond.WriteLog(ThisReceive, $"開啟Excel Application");
                xlApp.Visible = MainServer.globalCommond.Excel_Applaction_Show;
                string filenamepath = MyCommond.Path_Program + MyCommond.Path_Report;
                bool fileExist = File.Exists(filenamepath + NewBookFileName);
                MyCommond.WriteLog(ThisReceive, $"新檔案是否存在:{fileExist}");
                Excel.Workbook OperationWB;
                if (fileExist) { OperationWB = xlApp.Workbooks.Open(Filename: filenamepath + NewBookFileName, ReadOnly: false); }
                else { OperationWB = xlApp.Workbooks.Open(Filename: filenamepath + BookFileName, ReadOnly: false); }

                MyCommond.WriteLog(ThisReceive, $"開啟Excel第 {Convert.ToString(Workbook_Table)} 籤頁");
                Excel.Worksheet ws2 = (Excel.Worksheet)OperationWB.Worksheets[Workbook_Table];

                #region 月合併
                if (IsNextMonth)
                {
                    if (IsFirstYear) { RowOffset = MonthAgo - MainServer.globalCommond.Operation_First_Day.Month + 2; }
                    else if (YearAgo > Reserve_Year) { RowOffset = Reserve_Year + MonthAgo; }
                    else { RowOffset = YearAgo + MonthAgo; }

                    MyCommond.WriteLog(ThisReceive, $"變數RowOffset狀態        :{Convert.ToString(RowOffset)}");
                    int tgR = StartRow + RowOffset;
                    int tgC = 0;
                    string msgStr;

                    var breakColumn = (from p in S_R_F_Line group p by new { p.EndColumn } into g select (MyCommond.ExcelColumnNum(g.Key.EndColumn) + 1)).ToList();

                    MyCommond.WriteLog(ThisReceive, $"月合併開始");
                    while (tgC <= EndColumn)
                    {
                        MyCommond.WriteLog(ThisReceive, $"開始處理Column第 {tgC} 欄的資料");
                        if (tgC == ReportDateStartColumn)
                        {
                            MyCommond.WriteLog(ThisReceive, $"在Row第 {Convert.ToString(tgR)} 行前插入一行");
                            ws2.Rows[tgR].insert();
                            //取得要合併的欄位
                            strRange = MyCommond.ExcelColumnStr(tgC) + (tgR) + ":" + MyCommond.ExcelColumnStr(tgC + 1) + (tgR);
                            MyCommond.WriteLog(ThisReceive, $"指定合併範圍 {strRange}");
                            //合併欄位
                            MyCommond.WriteLog(ThisReceive, $"將該範圍合併、文字置中、將格式定義為YYYY.MM");
                            ws2.Range[strRange].Merge();
                            //將合併欄位的文字置中
                            ws2.Cells[tgR, tgC].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            //修改格式
                            ws2.Cells[tgR, tgC].NumberFormatLocal = "yyyy.mm";
                            //取得日期
                            var ddd = ws2.Cells[tgR + 1, tgC].text;
                            StrDate = Convert.ToDateTime(ws2.Cells[tgR + 1, tgC].text);
                            MyCommond.WriteLog(ThisReceive, $"取得第 {Convert.ToString(tgR + 1)} 行日期為 {StrDate.ToString()}");
                            //輸入營運日前的日期
                            MyCommond.WriteLog(ThisReceive, $"將日期寫入 {MyCommond.ExcelColumnStr(tgC)}{Convert.ToString(tgR)}");
                            ws2.Cells[tgR, tgC] = StrDate.ToString("yyyy/MM/01");
                        }   //寫入該月份
                        else if (tgC == ReportDateStartColumn + 1)
                        {
                            MyCommond.WriteLog(ThisReceive, $"此欄不做處理");
                        }   //星期
                        else if (tgC == ReportDateStartColumn + 2)
                        {
                            // =SUM(E12:J12,L12:Q12,S12:W12)
                            
                            string ForString = "";
                            foreach (var item in S_R_F_Line)
                            {
                                ForString += $"{((ForString != "") ? "," : "")}{item.StartColumn}{tgR}:{item.EndColumn}{tgR}";
                                string ForString2 = $@"=SUM({item.StartColumn}{tgR}:{item.EndColumn}{tgR})";
                                ws2.Cells[tgR, MyCommond.ExcelColumnNum(item.EndColumn) + 1] = ForString2;
                            }
                            msgStr = $"SUM({ForString})";

                            ws2.Cells[tgR, tgC] = string.Format("=SUM({0})",ForString);
                            MyCommond.WriteLog(ThisReceive, $"={msgStr}");
                        }   //總計
                        else if (tgC >= StartColumn && tgC <= EndColumn)
                        {
                            if (breakColumn.Contains(tgC))
                            {
                                

                            }
                            else
                            {
                                //設定範圍
                                strRange = MyCommond.ExcelColumnStr(tgC) + (tgR + 1) + ":" + MyCommond.ExcelColumnStr(tgC) + (tgR + LastMonthDayCount);
                                //取得要計算的範圍
                                myRange = ws2.Range[strRange];
                                //將範圍內的值做加總
                                SumValue = Convert.ToInt32(xlApp.WorksheetFunction.Sum(myRange));
                                msgStr = string.Format(@"將範圍 {0} 加總和為:{1} 寫入{2}{3} ,並清除範圍資料", strRange, SumValue, MyCommond.ExcelColumnStr(tgC), Convert.ToString(tgR));
                                MyCommond.WriteLog(ThisReceive, $"{msgStr}");
                                //將加總的值輸入
                                ws2.Cells[tgR, tgC] = SumValue.ToString();
                                //清除資料
                                myRange.ClearContents();
                            }
                        }   //所有票分別加總
                        tgC++;
                    }
                    //設定本月第一天日期
                    MyCommond.WriteLog(ThisReceive, $"在 {S_R_F_Line[0].ReportDateColumn}:{Convert.ToString(tgR + 1)} 寫入總和 {Convert.ToString(OperationDate.AddDays(-0).ToString("yyyy/MM/dd"))}");
                    ws2.Cells[tgR + 1, MyCommond.ExcelColumnNum(S_R_F_Line[0].ReportDateColumn)] = OperationDate.AddDays(-0).ToString("yyyy/MM/dd");
                    MyCommond.WriteLog(ThisReceive, $"月合併結束");
                }
                #endregion

                #region 年合併
                if (IsNextYear)
                {
                    string msgStr;
                    // if (YearAgo > Reserve_Year) { RowOffset = Reserve_Year; } else { RowOffset = YearAgo; }
                    MyCommond.WriteLog(ThisReceive, $"年合併開始");

                    int EndList = S_R_F_Line.Max(x => MyCommond.ExcelColumnNum(x.EndColumn));

                    int tgR = StartRow;
                    int tgC = ReportDateStartColumn;
                    msgStr = "將{0}:{1}寫到{0}:{2}";

                    ValueCount += 1;

                    //填寫或搬移年總運量(保留近3年)
                    while (tgC <= EndList)
                    {
                        if (tgC == ReportDateStartColumn)
                        {
                            if (YearAgo <= Reserve_Year)
                            {
                                MyCommond.WriteLog(ThisReceive, $"在第 {Convert.ToString(tgR)}前插入一行");
                                ws2.Rows[tgR].insert();
                                //取得要合併的欄位
                                strRange = MyCommond.ExcelColumnStr(tgC) + (tgR) + ":" + MyCommond.ExcelColumnStr(tgC + 1) + (tgR);
                                //合併欄位
                                ws2.Range[strRange].Merge();
                                //將合併欄位的文字置中
                                ws2.Cells[tgR, tgC].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                //取得日期
                                StrDate = Convert.ToDateTime(ws2.Cells[tgR + 1, tgC].text);
                                //將日期重新編排輸入
                                ws2.Cells[tgR, tgC].NumberFormatLocal = "yyyy年";
                                ws2.Cells[tgR, tgC] = StrDate.ToString("yyyy/01/01");
                                //msgStr = "取得 " + MyCommond.ExcelColumnStr(tgC) + ":" + Convert.ToString(tgR + 1) + " 資料 " + StrDate + " 寫入 " + MyCommond.ExcelColumnStr(tgC) + Convert.ToString(tgR) + "並將設改為YYYY.MM";
                                msgStr = string.Format("取得{0}:{1}資料{2}寫入{3}{4}，並將資料型態設為yyyy.MM", MyCommond.ExcelColumnStr(tgC), Convert.ToString(tgR + 1), StrDate, MyCommond.ExcelColumnStr(tgC), Convert.ToString(tgR));
                                MyCommond.WriteLog(ThisReceive, $"{msgStr}");
                            }
                            else
                            {
                                ws2.Cells[tgR, ReportDateStartColumn] = ws2.Cells[tgR + 1, ReportDateStartColumn];
                                //ws2.Cells[tgR - 1, tgC] = ws2.Cells[tgR, tgC];
                                MyCommond.WriteLog(ThisReceive, $"將cells({tgR + 1},{ReportDateStartColumn})資料放到cells({tgR},{ReportDateStartColumn})");
                            }
                        }   //寫入該月份
                        else if (tgC == ReportDateStartColumn + 1)
                        {
                            MyCommond.WriteLog(ThisReceive, $"此欄不做處理");
                        }   //星期
                        else if (tgC == ReportDateStartColumn + 2)
                        {
                            // =SUM(E12:J12,L12:Q12,S12:W12)

                            string ForString = "";
                            foreach (var item in S_R_F_Line)
                            {
                                ForString += $"{((ForString != "") ? "," : "")}{item.StartColumn}{tgR}:{item.EndColumn}{tgR}";
                                string ForString2 = $@"=SUM({item.StartColumn}{tgR}:{item.EndColumn}{tgR})";
                                ws2.Cells[tgR, MyCommond.ExcelColumnNum(item.EndColumn) + 1] = ForString2;
                            }
                            msgStr = $"SUM({ForString})";
                            ws2.Cells[tgR, tgC] = string.Format("=SUM({0})", ForString);
                            MyCommond.WriteLog(ThisReceive, $"={msgStr}");
                        }   //總計
                        else if (tgC >= StartColumn && tgC <= EndList)
                        {
                            if (YearAgo > 3)
                            {
                                ws2.Cells[tgR - 2, tgC] = ws2.Cells[tgR - 1, tgC];
                                ws2.Cells[tgR - 1, tgC] = ws2.Cells[tgR, tgC];
                            }
                            //設定範圍
                            strRange = MyCommond.ExcelColumnStr(tgC) + (tgR + 1) + ":" + MyCommond.ExcelColumnStr(tgC) + (tgR + MonthAgo);
                            //取得要計算的範圍
                            myRange = ws2.Range[strRange];
                            //將範圍內的值做加總
                            SumValue = Convert.ToInt32(xlApp.WorksheetFunction.Sum(myRange));
                            msgStr = string.Format("將範圍 {0} 加總，和為:{1} 寫入{2}{3}", strRange, SumValue, MyCommond.ExcelColumnStr(tgC), Convert.ToString(tgR));
                            MyCommond.WriteLog(ThisReceive, $"msgStr");
                            //將加總的值輸入
                            ws2.Cells[tgR, tgC] = SumValue.ToString();
                        }   //所有票分別加總
                        tgC++;
                    }

                    ValueCount += 1;

                    MyCommond.WriteLog(ThisReceive, $"開始移除每月運量");
                    //移除每月總運量
                    Int32 WrithTime = 0;
                    strRange = (tgR + 1) + ":" + (tgR + MonthAgo);
                    ws2.Rows[strRange].delete();
                }
                #endregion

                #region 合併資料
                MyCommond.WriteLog(ThisReceive, $"合併空間中的資料");
                var Line1 = MainServer.globalCommond.Stock_ReportFile.Find(x =>
                    x.ReportNameEng == MethodeName_String &&
                    x.OperationLine == MainServer.globalCommond.OperationCode);

                var Merge = Data_Afc;

                var ReportStart = MyCommond.ExcelColumnNum(Line1.StartColumn);
                var ReportMobile = MyCommond.ExcelColumnNum(Line1.MobilePayColumn);
                var ReportOwn = MyCommond.ExcelColumnNum(Line1.OwnTiketColumn);
                try
                {
                    for (int i = 0; i < Data_MobilePay.GetLength(0); i++)
                    {
                        for (int j = 0; j < Data_MobilePay.GetLength(1); j++)
                        {
                            Merge[0 + i, ReportMobile - ReportStart + j] += Data_MobilePay[i, j];
                        }
                    }
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
                }

                try
                {
                    for (int i = 0; i < Data_OwnTicket.GetLength(0); i++)
                    {
                        for (int j = 0; j < Data_OwnTicket.GetLength(1); j++)
                        {
                            Merge[0 + i, ReportOwn - ReportStart + j] += Data_OwnTicket[i, j];
                        }
                    }
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
                }

                #endregion 合併資料

                #region 輸出

                MyCommond.WriteLog(ThisReceive, $"開始寫入每日資料");
                Int32 WriteRow, WriteColumn;

                IsFirstYear = OperationDate.AddDays(-0).Year == MainServer.globalCommond.Operation_First_Day.Year;
                MyCommond.WriteLog(ThisReceive, $"變更IsFirstYear狀態      :{Convert.ToString(IsFirstYear)}");

                MonthAgo = Convert.ToDateTime(OperationDate.ToString("yyyy/MM/01")).AddMonths(-1).Month;
                if (MonthAgo == 12) { MonthAgo = 0; }

                if (IsFirstYear) MonthAgo = MonthAgo - MainServer.globalCommond.Operation_First_Day.Month;
                MyCommond.WriteLog(ThisReceive, $"變數MonthAgo狀態:{Convert.ToString(MonthAgo)}");

                ValueCount += 1;

                WriteRow = StartRow + MonthAgo + ThisDay - 1;
                if (!IsFirstYear) WriteRow += YearAgo;

                MyCommond.WriteLog(ThisReceive, $"寫入起始行:{WriteRow}");

                ValueCount += 1;

                string msgStr_Y = "第 {0} 行第 {1} 欄寫入第 {2} 筆資料：{3}";
                WriteColumn = MyCommond.ExcelColumnNum(Line1.StartColumn);
                MyCommond.WriteLog(ThisReceive, $"寫入起始欄:{WriteColumn}");
                int sss = 0;
                for (int i = 0; i < Merge.GetLength(0); i++)
                {
                    for (int ColumnOffset2 = 0; ColumnOffset2 < Merge.GetLength(1); ColumnOffset2++)
                    {
                        int s1 = WriteRow;
                        int s2 = WriteColumn + ColumnOffset2;
                        int s3 = (i + 1) * (ColumnOffset2 + 1);
                        int s4 = Merge[i, ColumnOffset2];
                        MyCommond.WriteLog(ThisReceive, $"{string.Format(msgStr_Y, s1, s2, sss++, s4)}");
                        ws2.Cells[s1, s2] = s4;
                    }
                }
                MyCommond.WriteLog(ThisReceive, $"");

                ValueCount += 1;

                #endregion 輸出

                //設定儲存檔名
                if (fileExist)
                {
                    MyCommond.WriteLog(ThisReceive, $"儲存檔案");
                    OperationWB.Save();
                }
                else
                {
                    MyCommond.WriteLog(ThisReceive, $"另存成:{NewBookFileName}");
                    ws2.SaveAs(filenamepath + NewBookFileName);
                }

                MyCommond.WriteLog(ThisReceive, $"關閉檔案");
                OperationWB.Close();
                MyCommond.WriteLog(ThisReceive, $"釋放Excel物件");
                Marshal.ReleaseComObject(ws2);
                Marshal.ReleaseComObject(OperationWB);
                ExportCheck = true;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{MyCommond.ExportErrorMessageToLog(ex, MethodeName.ToDescription()).Message}");
                ExportCheck = false;
            }
            finally
            {
                MyCommond.WriteLog(ThisReceive, $"關閉Excel Application");
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                MainServer.VA_Lock = false;
            }

            ValueCount += 1;

            #endregion 運量、營收

            MyCommond.WriteLog(ThisReceive, $"客製化運量、營收輸出 - 結束");
            return ExportCheck;

        NormalSub:
            MyCommond.WriteLog(ThisReceive, $"其他通用輸出 - 開始");

            #region 普通

            ValueCount += 1;

            bool ckeck = MyCommond.CheckFile(MyCommond.Path_Program + MyCommond.Path_Report + ExportFileName);
            XLWorkbook wb = new XLWorkbook();
            if (ckeck)
            {
                MyCommond.WriteLog(ThisReceive, "以既有的檔案複寫");
                wb = new XLWorkbook(MyCommond.Path_Program + MyCommond.Path_Report + ExportFileName);
            }
            else
            {
                if (!(MyCommond.CheckFile(template)))
                {
                    MyCommond.WriteLog(ThisReceive, "無樣版檔，認後再重新執行");

                    Thread thread = new Thread(() => MessageBox.Show($"無樣版檔\n{template}\n認後再重新執行", "錯誤"));
                    thread.IsBackground = true;
                    thread.Start();
                    return ExportCheck;
                }

                MyCommond.WriteLog(ThisReceive, "以樣板檔案寫入");
                wb = new XLWorkbook(template);
            }

            ValueCount += 1;

            try
            {
                var ws = wb.Worksheets.First();
                int RowBreak, ColumnBreak;

                MyCommond.WriteLog(ThisReceive, "寫入產製日期");
                if (ReportDateStartRow > 0 && ReportDateStartColumn > 0)
                {
                    ws.Cell(ReportDateStartRow, ReportDateStartColumn).Value = $"{OperationDate}";
                    if (!(MethodeName == LrtTxnClientSwitch.Month_Volume || MethodeName == LrtTxnClientSwitch.Month_Amount))
                    {
                        ws.Cell(ReportDateStartRow + 1, ReportDateStartColumn).Value = DateTime.Now.ToString("yyyy/MM/dd");
                    }
                }

                ValueCount += 1;

                #region 一般資料寫入

                try
                {
                    MyCommond.WriteLog(ThisReceive, $"Data_Afc -> Row:{Data_Afc.GetLength(0),4}, Column:{Data_Afc.GetLength(1),4}");
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"Data_Afc -> Error\n{ex.Message}");
                }
                finally
                {

                }

                if (Data_Afc != null && Data_Afc.GetLength(0) > 0 && Data_Afc.GetLength(1) > 0)
                {
                    if (thisBreakOffRow != 0)
                    {
                        RowBreak = thisBreakOffRow - StartRow + 1;
                        EndRow = EndRow - StartRow - thisBreakOffRowNum;
                    }
                    else
                    {
                        RowBreak = Data_Afc.GetLength(0);
                    }
                    if (thisBreakOffColumn != 0)
                    {
                        ColumnBreak = thisBreakOffColumn - StartColumn + 1;
                        EndColumn = EndColumn - StartColumn - thisBreakOffColumnNum + 1;
                    }
                    else
                    {
                        ColumnBreak = Data_Afc.GetLength(1);
                    }

                    MyCommond.WriteLog(ThisReceive, $"---RowBreak:{RowBreak,4}, EndRow:{EndRow,4}, ColumnBreak:{ColumnBreak,4}, EndColumn:{EndColumn,4}---");

                    ValueCount += 1;

                    MyCommond.WriteLog(ThisReceive, "開始資料寫入");
                    for (int RowOffset_1 = 0; RowOffset_1 < RowBreak; RowOffset_1++)
                    {
                        try
                        {
                            for (int ColumnOffset = 0; ColumnOffset < ColumnBreak; ColumnOffset++)
                            {
                                MyCommond.WriteLog(ThisReceive, $"" +
                                    $"sheet.Cell({(StartRow + RowOffset_1),4}, {(StartColumn + ColumnOffset),4}).Value = " +
                                    $"{Data_Afc[RowOffset_1, ColumnOffset]}");
                                ws.Cell(StartRow + RowOffset_1, StartColumn + ColumnOffset).Value = Data_Afc[RowOffset_1, ColumnOffset];
                            }
                            if (thisBreakOffColumn != 0)
                            {
                                for (int ColumnOffset = ColumnBreak; ColumnOffset < EndColumn; ColumnOffset++)
                                {
                                    MyCommond.WriteLog(ThisReceive, $"" +
                                        $"sheet.Cell({(StartRow + RowOffset_1),4}, {(StartColumn + thisBreakOffColumnNum + ColumnOffset),4}).Value = " +
                                        $"{Data_Afc[RowOffset_1, ColumnOffset]}");
                                    ws.Cell(StartRow + RowOffset_1, StartColumn + thisBreakOffColumnNum + ColumnOffset).Value = Data_Afc[RowOffset_1, ColumnOffset];
                                }
                            }  //斷點欄後寫數值不對
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"ERROR! \n{MyCommond.ExportErrorMessageToLog(ex, MethodeName.ToDescription()).Message}");
                        }
                    }

                    MyCommond.WriteLog(ThisReceive, "行中斷資料寫入");
                    if (thisBreakOffRow != 0)
                    {
                        for (int RowOffset_2 = RowBreak; RowOffset_2 < EndRow; RowOffset_2++)
                        {
                            try
                            {
                                for (int ColumnOffset = 0; ColumnOffset < ColumnBreak; ColumnOffset++)
                                {
                                    MyCommond.WriteLog(ThisReceive, $"" +
                                        $"sheet.Cell({(StartRow + thisBreakOffRowNum + RowOffset_2),4}, {(StartColumn + ColumnOffset),4}).Value = " +
                                        $"{Data_Afc[RowOffset_2, ColumnOffset]}");
                                    ws.Cell(StartRow + thisBreakOffRowNum + RowOffset_2, StartColumn + ColumnOffset).Value = Data_Afc[RowOffset_2, ColumnOffset];
                                }
                                if (thisBreakOffColumn != 0)
                                {
                                    for (int ColumnOffset = ColumnBreak; ColumnOffset < EndColumn; ColumnOffset++)
                                    {
                                        MyCommond.WriteLog(ThisReceive, $"" +
                                            $"sheet.Cell({(StartRow + thisBreakOffRowNum + RowOffset_2),4}, {(StartColumn + thisBreakOffColumnNum + ColumnOffset),4}).Value = " +
                                            $"{Data_Afc[RowOffset_2, ColumnOffset]}");
                                        ws.Cell(StartRow + thisBreakOffRowNum + RowOffset_2, StartColumn + thisBreakOffColumnNum + ColumnOffset).Value = Data_Afc[RowOffset_2, ColumnOffset];
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                MyCommond.WriteLog(ThisReceive, $"ERROR! \n{MyCommond.ExportErrorMessageToLog(ex, MethodeName.ToDescription()).Message}");
                            }
                        }
                    }
                }

                ValueCount += 1;

                #endregion

                #region 行動支付寫入

                try
                {
                    MyCommond.WriteLog(ThisReceive, $"Data_MobilePay -> Row:{Data_MobilePay.GetLength(0),4}, Column:{Data_MobilePay.GetLength(1),4}");
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"Data_MobilePay -> Error\n{ex.Message}");
                }
                finally
                {

                }

                if (Data_MobilePay != null && Data_MobilePay.GetLength(0) > 0 && Data_MobilePay.GetLength(1) > 0)
                {
                    MyCommond.WriteLog(ThisReceive, "行動支付寫入");
                    for (int RowOffset_3 = 0; RowOffset_3 < Data_MobilePay.GetLength(0); RowOffset_3++)
                    {
                        for (int ColumnOffset_3 = 0; ColumnOffset_3 < Data_MobilePay.GetLength(1); ColumnOffset_3++)
                        {
                            int RowSet = StartRow + RowOffset_3;
                            int ColumnSet = MobilePayColumn + ColumnOffset_3;
                            //int ExcelCellValue = Convert.ToInt32(Convert.ToString(ws.Cell(RowSet, ColumnSet).Value) ?? "0");
                            int ExcelCellValue = ((ws.Cell(RowSet, ColumnSet).Value == "") ? 0 : Convert.ToInt32(ws.Cell(RowSet, ColumnSet).Value));
                            MyCommond.WriteLog(ThisReceive, $"" +
                                $"sheet.Cell({RowSet}, {ColumnSet}).Value = " +
                                $"{ExcelCellValue} + {Data_MobilePay[RowOffset_3, ColumnOffset_3]} = " +
                                $"{ExcelCellValue + Data_MobilePay[RowOffset_3, ColumnOffset_3]}");
                            ws.Cell(RowSet, ColumnSet).Value = Convert.ToInt64(ExcelCellValue) + Data_MobilePay[RowOffset_3, ColumnOffset_3];
                        }
                    }
                }

                ValueCount += 1;

                #endregion

                #region 人工販售寫入

                try
                {
                    MyCommond.WriteLog(ThisReceive, $"Data_OwnTicket -> Row:{Data_OwnTicket.GetLength(0),4}, Column:{Data_OwnTicket.GetLength(1),4}");
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"Data_OwnTicket -> Error\n{ex.Message}");
                }
                finally
                {

                }

                if (Data_OwnTicket != null && Data_OwnTicket.GetLength(0) > 0 && Data_OwnTicket.GetLength(1) > 0)
                {
                    MyCommond.WriteLog(ThisReceive, $"自有票寫入");
                    for (int RowOffset_3 = 0; RowOffset_3 < Data_OwnTicket.GetLength(0); RowOffset_3++)
                    {
                        for (int ColumnOffset_3 = 0; ColumnOffset_3 < Data_OwnTicket.GetLength(1); ColumnOffset_3++)
                        {
                            if (Data_OwnTicket[RowOffset_3, ColumnOffset_3] == null) continue;
                            int RowSet = StartRow + RowOffset_3;
                            int ColumnSet = OwnTiketColumn + ColumnOffset_3;
                            MyCommond.WriteLog(ThisReceive, $"" +
                                $"sheet.Cell({RowSet}, {ColumnSet}).Value = " +
                                $"{((ws.Cell(RowSet, ColumnSet).Value != "") ? Convert.ToInt64(ws.Cell(RowSet, ColumnSet).Value) : 0)} + {Data_OwnTicket[RowOffset_3, ColumnOffset_3]} = " +
                                $"{((ws.Cell(RowSet, ColumnSet).Value != "") ? Convert.ToInt64(ws.Cell(RowSet, ColumnSet).Value) : 0) + Data_OwnTicket[RowOffset_3, ColumnOffset_3]}");
                            ws.Cell(RowSet, ColumnSet).Value = ((ws.Cell(RowSet, ColumnSet).Value != "") ? Convert.ToInt64(ws.Cell(RowSet, ColumnSet).Value) : 0) + Data_OwnTicket[RowOffset_3, ColumnOffset_3];
                        }
                    }
                }

                ValueCount += 1;

                #endregion

                MyCommond.WriteLog(ThisReceive, $"寫入結束，將檔案存成{ExportFileName}");
                ExportCheck = true;
                ValueCount += 1;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR! \n{MyCommond.ExportErrorMessageToLog(ex, MethodeName.ToDescription()).Message}");
            }
            finally
            {
                wb.SaveAs(ExportPath + ExportFileName);
                wb = null;
                MyCommond.WriteLog(ThisReceive, $"結束Excel輸出");
            }

            #endregion

            MyCommond.WriteLog(ThisReceive, $"其他通用輸出 - 結束");

            return ExportCheck;
        }

        #endregion

        #endregion

    }
}
