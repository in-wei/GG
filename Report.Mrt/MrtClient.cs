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
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using GlobalCommond;
using GlobalCommond.ViewModel;
using ClosedXML.Excel;

namespace Report.Mrt
{
    public class MrtClient
    {
        #region 設定

        public Guid ID { get; private set; }
        public MrtMain MainServer { get; private set; }
        public MrtTxnClientSwitch MrtMethodeName { get; private set; }
        public StatusEnum Status { get; private set; }

        public string CardSN { get; private set; }
        public int Counts { get; private set; }
        public List<string> CardList { get; private set; }
        public List<TxnTransaction> CardTxnData { get; private set; }

        #endregion

        #region 變數

        private MyCommonds MyCommond { get; set; }
        private string ThisReceive { get; set; }
        private Thread sendingThread { get; set; }

        List<int> mStation_0 = new List<int> { 200, 201, 202, 203, 204, 205, 206, 207, 208, 209, 210, 211, 212, 213 };
        List<int> mStation_1 = new List<int> { 201, 202, 203, 205, 206, 207, 208, 212, 213 };
        List<string> mTxnType = new List<string> { "EO", "TO" };
        TimeSpan OverTime = new TimeSpan(0, 30, 0);

        Thread threadCalculator_Analyze { get; set; }
        List<TxnTransaction> ResponseExportTxnNormal { get; set; }
        List<TxnTransaction> ResponseExportTxnMeger { get; set; }
        List<TxnTransaction> ResponseExportTxnError { get; set; }
        List<TxnTransaction> ResponseExportTxnCompare { get; set; }

        #endregion

        public MrtClient()
        {
            MyCommond = new MyCommonds();
            //MyCommond = MainServer.globalCommond.MyCommond;
            CardTxnData = new List<TxnTransaction>();
            ResponseExportTxnNormal = new List<TxnTransaction>();
            ResponseExportTxnCompare = new List<TxnTransaction>();
            ResponseExportTxnMeger = new List<TxnTransaction>();
            ResponseExportTxnError = new List<TxnTransaction>();
            CardList = new List<string>();
            Counts = 0;
            Status = StatusEnum.Disconnected;
        }

        public MrtClient(MrtMain GetMainServer, MrtTxnClientSwitch mMethodeName) : this()
        {
            MainServer = GetMainServer;
            MrtMethodeName = mMethodeName;
            ID = Guid.NewGuid();
            ThisReceive = $@"client\({ID.ToString()})";
            Status = StatusEnum.Ready;
            MyCommond.WriteLog(ThisReceive, "創建新Clien端");
        }

        #region 功能

        public void Start()
        {

            //receivingThread = new Thread(ReceivingMethod);
            //receivingThread.IsBackground = true;
            //receivingThread.Start();

            sendingThread = new Thread(SendingMethod);
            sendingThread.IsBackground = true;
            sendingThread.Start();

        }

        public void Next()
        {
            MrtMethodeName = MrtTxnClientSwitch.NA;
            CardSN = "";
            Counts = 0;
            CardTxnData.Clear();
        }

        #endregion

        #region 非同步線呈執行程式

        private void SendingMethod(object obj)
        {
            MyCommond.WriteLog(ThisReceive, $"開始計算 {MrtMethodeName.ToString()}");
            while (Status != StatusEnum.Disconnected)
            {
                if (Status == StatusEnum.Ready)
                {
                         if (MrtMethodeName == MrtTxnClientSwitch.NA)
                    {
                        // Null
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.EzsyAnalyze)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_EasyAnalyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.TxnTypeAnalyze)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_TxnTypeAnalyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.IOAnalyze)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_IOAnalyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.CompareAnalyze)
                    {
                        try
                        {
                            if (true)
                            {
                                Status = StatusEnum.Calculating;
                                threadCalculator_Analyze = new Thread(Calculator_CompareAnalyze);
                                threadCalculator_Analyze.IsBackground = true;
                                threadCalculator_Analyze.Start();
                            }
                            else
                            {
                                if (MainServer.TxnCard_more3.Count() <= 0) continue;
                                if (MainServer.CardCount_More >= MainServer.TxnCard_more3.Count())
                                {
                                    Status = StatusEnum.Disconnected;
                                    continue;
                                }
                                CardSN = MainServer.GetCardSN(ID.ToString());
                                CardTxnData = MainServer.GetCardTxnData(ID.ToString(), CardSN);
                                Counts++;
                                Status = StatusEnum.Calculating;
                                threadCalculator_Analyze = new Thread(Calculator_CompareAnalyze);
                                threadCalculator_Analyze.IsBackground = true;
                                threadCalculator_Analyze.Start();
                                if (!threadCalculator_Analyze.IsAlive)
                                {
                                    threadCalculator_Analyze = new Thread(ExportTxn_CompareAnalyze);
                                    threadCalculator_Analyze.IsBackground = true;
                                    threadCalculator_Analyze.Start();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.TrnAnalyze)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_TrnAnalyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.TPassAnalyze)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_TPassAnalyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.VolumnReportAnalyze)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_VolumeReportAnalyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.Day_AddValue)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_DayAddValueAnalyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.Day_EveryIssuer)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_EveryIssuerAnalyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.C10Analyze)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_C10Analyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.C11Analyze)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_C11Analyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.C12Analyze)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_C12Analyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.C13Analyze)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_C13Analyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.C14Analyze)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_C14Analyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.C15Analyze)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_C15Analyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.C16Analyze)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_C16Analyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.C17Analyze)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_C17Analyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.C18Analyze)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_C18Analyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.C19Analyze)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_C19Analyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                    else if (MrtMethodeName == MrtTxnClientSwitch.C20Analyze)
                    {
                        try
                        {
                            Status = StatusEnum.Calculating;
                            threadCalculator_Analyze = new Thread(Calculator_C20Analyze);
                            threadCalculator_Analyze.IsBackground = true;
                            threadCalculator_Analyze.Start();
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                        }
                    }
                }
                Thread.Sleep(30);
            }
        }

        private void ReceivingMethod(object obj)
        {
            while (Status != StatusEnum.Disconnected)
            {
                //if (TcpClient.Available > 0)
                //{
                //    //try
                //    //{
                //    BinaryFormatter f = new BinaryFormatter();
                //    f.Binder = new Shared.AllowAllAssemblyVersionsDeserializationBinder();
                //    MessageBase msg = f.Deserialize(TcpClient.GetStream()) as MessageBase;
                //    OnMessageReceived(msg);
                //    //}
                //    //catch (Exception e)
                //    //{
                //    //    Exception ex = new Exception("Unknown message recieved. Could not deserialize the stream.", e);
                //    //    OnClientError(this, ex);
                //    //    Debug.WriteLine(ex.Message);
                //    //}
                //}

                Thread.Sleep(30);
            }
        }

        #region 已使用

        #region 簡易資料(每日)

        private void Calculator_EasyAnalyze()
        {
            //MyCommond.WriteLog(ThisReceive, $"開始計算 簡易資料(每日)");
            Status = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"簡易資料計算 - 開始");
            string ExportTxnTransactionMegreTitle = (new TxnTransactionMegre()).ExportTitle_English();
            try
            {
                #region 所有交易類別分類總計

                CardSN = MainServer.OperationDate.ToString("yyyyMMdd");
                DateTime StartTime = DateTime.Now;
                string FileName = $"history{MainServer.OperationDate:yyyyMMdd}_nc.csv";
                string[] Data = null;
                int CountExit, ExitAmt, RideCountAdd = 0, RideCountMinus = 0, AmtAdd = 0, AmtMinus = 0;
                Int32 CountData = 0, CountRun = 0;
                TimeSpan CTime;

                if (true)
                {
                    int Pass1280Ride = 0, Pass1200Ride = 0, Pass1200Set = 0;
                    int Pass1200RefundCount = 0, Pass1200RefundAmt = 0;
                    int Pass1200CancelCount = 0, Pass1200CancelAmt = 0;
                    bool TxnTypeSubCheck = false;

                    List<TxnTransaction> StockMrtAllTxn = (
                        from p in MainServer.TxnAllData
                        where
                            (
                                MainServer.globalCommond.Stock_Mrt_Y_StationOd.FindIndex(x => x.ExitStation == p.SVCE_LOC_ID && x.IsCount == "Y") != -1
                                || mStation_0.Contains(Convert.ToInt16(p.SVCE_LOC_ID))
                            )
                            && Convert.ToDateTime(p.TXN_TIMESTAMP) >= MainServer.OperationDate
                            && Convert.ToDateTime(p.TXN_TIMESTAMP) < MainServer.OperationDate.AddDays(1)
                        select p
                    ).ToList();
                    MyCommond.WriteLog(ThisReceive, $"總交易資料篩選後比數:{StockMrtAllTxn.Count}");

                    List<int> ls_O = new List<int>();
                    foreach (var item in MainServer.globalCommond.Stock_Mrt_TxnType)
                    {
                        if (item.mCheck != null)
                        {
                            if (item.mCheck.Length == 2 && item.mCheck.Substring(1, 1) == "O")
                            {
                                ls_O.Add(item.Num);
                            }
                            else if (item.mCheck.Length == 1 && item.mCheck == "T")
                            {
                                // ls_O.Add(item.Num);
                            }
                        }
                    }

                    List<int> ls_R = new List<int>();
                    foreach (var item in MainServer.globalCommond.Stock_Mrt_TxnType)
                    {
                        if (item.mCheck != null && item.mCheck.Length == 2)
                        {
                            if (item.mCheck.Length == 2 && item.mCheck.Substring(0,1) == "T" && (item.mCheck.Substring(1, 1) == "R" || item.mCheck.Substring(1, 1) == "F"))
                            {
                                ls_R.Add(item.Num);
                            }
                        }
                    }

                    List<TxnTransaction> RT_IS_ECC_1 = (from p in StockMrtAllTxn where ls_O.Contains(Convert.ToInt16(p.CARD_TXN_TYPE_ID)) && p.ISSUER_ID == "255" select p).ToList();
                    List<TxnTransaction> RT_IS_IPASS_1 = (from p in StockMrtAllTxn where ls_O.Contains(Convert.ToInt16(p.CARD_TXN_TYPE_ID)) && p.ISSUER_ID == "9" select p).ToList();
                    List<TxnTransaction> RT_IS_ICASH_1 = (from p in StockMrtAllTxn where ls_O.Contains(Convert.ToInt16(p.CARD_TXN_TYPE_ID)) && p.ISSUER_ID == "11" select p).ToList();
                    MyCommond.WriteLog(ThisReceive, $"交易資料(ECC):{RT_IS_ECC_1.Count}");
                    MyCommond.WriteLog(ThisReceive, $"交易資料(IPAS):{RT_IS_IPASS_1.Count}");
                    MyCommond.WriteLog(ThisReceive, $"交易資料(ICASH):{RT_IS_ICASH_1.Count}");

                    //MyCommond.WriteLog(ThisReceive, $"初始化變數(RT_O)");
                    List<TxnTransaction> RT_O = new List<TxnTransaction>();
                    //MyCommond.WriteLog(ThisReceive, $"在變數(RT_O)加入ECC");
                    RT_O.AddRange((from p in RT_IS_ECC_1   where MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ecc.Find(x => Convert.ToString(x.Num) == p.CARD_TXN_SUBTYPE_ID).mCheck == "o" select p).ToList());
                    //MyCommond.WriteLog(ThisReceive, $"在變數(RT_O)加入IPASS");
                    RT_O.AddRange((from p in RT_IS_IPASS_1 where MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ipass.Find(x => Convert.ToString(x.Num) == p.CARD_TXN_SUBTYPE_ID).mCheck == "o" select p).ToList());
                    //MyCommond.WriteLog(ThisReceive, $"在變數(RT_O)加入ICASH");
                    RT_O.AddRange((from p in RT_IS_ICASH_1 where MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Icash.Find(x => Convert.ToString(x.Num) == p.CARD_TXN_SUBTYPE_ID).mCheck == "o" select p).ToList());
                    MyCommond.WriteLog(ThisReceive, $"計算交易資料總比數:{RT_O.Count}");

                    List <TxnTransaction> RT_R = (
                        from p in StockMrtAllTxn
                        where ls_R.Contains(Convert.ToInt16(p.CARD_TXN_TYPE_ID))
                        select p
                        ).ToList();

                    List<TxnTransaction> RT_SS = (
                        from p in StockMrtAllTxn
                        where Convert.ToInt16(p.CARD_TXN_TYPE_ID) == 3 && Convert.ToInt16(p.TICKET_SUBTYPE_ID) == 75
                        select p
                        ).ToList();
                    MyCommond.WriteLog(ThisReceive, $"行政院通勤月票_設定:{RT_SS.Count}");

                    List<TxnTransaction> RT_SR = (
                        from p in StockMrtAllTxn
                        where Convert.ToInt16(p.CARD_TXN_TYPE_ID) == 6 && Convert.ToInt16(p.TICKET_SUBTYPE_ID) == 75
                        select p
                        ).ToList();
                    MyCommond.WriteLog(ThisReceive, $"行政院通勤月票_退票:{RT_SR.Count}");

                    List<TxnTransaction> RT_SC = (
                        from p in StockMrtAllTxn
                        where Convert.ToInt16(p.CARD_TXN_TYPE_ID) == 10 && Convert.ToInt16(p.TICKET_SUBTYPE_ID) == 75
                        select p
                        ).ToList();
                    MyCommond.WriteLog(ThisReceive, $"aaaaa");

                    var t1 = RT_O.FindAll(x => x.FARE_PRODUCT_TYPE_ID == "74").ToList();
                    var t2 = RT_O.FindAll(x => x.FARE_PRODUCT_TYPE_ID == "75").ToList();
                    var t3 = RT_R.FindAll(x => x.FARE_PRODUCT_TYPE_ID == "74").ToList();
                    var t4 = RT_R.FindAll(x => x.FARE_PRODUCT_TYPE_ID == "75").ToList();

                    RideCountAdd = RT_O.Count;
                    AmtAdd = RT_O.Sum(x => Convert.ToInt32(x.TXN_AMT));
                    RideCountMinus = RT_R.Count;
                    AmtMinus = RT_R.Sum(x => Convert.ToInt32(x.TXN_AMT));
                    Pass1280Ride = (t1.Count) - (t3.Count);
                    Pass1200Ride = (t2.Count) - (t4.Count);
                    Pass1200Set = RT_SS.Count;
                    Pass1200RefundCount = RT_SR.Count;
                    Pass1200RefundAmt = RT_SR.Sum(x => Convert.ToInt32(x.TXN_AMT));
                    Pass1200CancelCount = RT_SC.Count;
                    Pass1200CancelAmt = RT_SC.Sum(x => Convert.ToInt32(x.TXN_AMT));

                    string mPath_ReportFile = MyCommond.Path_Program + MyCommond.Path_Report + $@"{MainServer.globalCommond.OperationCode_String}\" + string.Format(MyCommond.Path_Date, MainServer.OperationDate.ToString("yyyy.MM.dd"));
                    string ReportfileName = mPath_ReportFile + MyCommond.CheckAndReName(mPath_ReportFile, $"行政院通勤月票_{MainServer.globalCommond.OperationCode_String}_{MainServer.OperationDate:yyyyMMdd}", "csv");
                    string data = $"日期,路線,1280月票進站,1280月票出站,1200月票進站,1200月票出站,1200月票設定,1200月票取消,1200月票退票,1200月票退票金額,1200退票手續費,1200總設定筆數" + Environment.NewLine +
                        $"{MainServer.OperationDate:yyyyMMdd},{MainServer.globalCommond.OperationCode},N,{Pass1280Ride},N,{Pass1200Ride},{Pass1200Set},{Pass1200CancelCount},{Pass1200RefundCount},{Pass1200RefundAmt},N,{Pass1200Set - Pass1200CancelCount}";
                    //WriteLog(ExecutionMode.Simple, MethodBase.GetCurrentMethod().Name.ToString(), Environment.NewLine + data);
                    MyCommond.WriteLog(ThisReceive, $"FullName:{ReportfileName}");
                    MyCommond.WriteLog(ThisReceive, $"Data    :{data}");
                    MyCommond.ExportData(ReportfileName, Encoding.UTF8, data);

                    DateTime EndTime = DateTime.Now;
                    CTime = EndTime - StartTime;
                    CountExit = RideCountAdd - RideCountMinus;
                    ExitAmt = AmtAdd - AmtMinus;
                    Status = StatusEnum.Finish;
                    return;
                    ///MessageBox.Show($"計算時間:{CTime}\n" +
                    ///    $"運量 => 正數:{RideCountAdd,8}, 減數:{RideCountMinus,8}, 總計:{CountExit,8}\n" +
                    ///    $"金額 => 正數:{AmtAdd,8}, 減數:{AmtMinus,8}, 總計:{ExitAmt,8}");

                }

                string mTidyfileName = MyCommond.Path_Program + MyCommond.Path_Tidy + string.Format(MyCommond.Path_Date, MainServer.OperationDate.ToString("yyyy.MM"));
                string TidyfileName = mTidyfileName + MyCommond.CheckAndReName(mTidyfileName, $"EveryData_{MainServer.OperationDate:yyyyMMdd}", "csv");
                if (!true)
                {
                    //WriteLog(ExecutionMode.Simple, MethodBase.GetCurrentMethod().Name.ToString(), "基礎運量計算開始");
                    //WriteLog(ExecutionMode.Simple, MethodBase.GetCurrentMethod().Name.ToString(), $"資料分析後儲存為 {TidyfileName}");

                    //List<TxnTransaction> mTransaction = new List<TxnTransaction>();

                    //WriteLog(ExecutionMode.Simple, MethodBase.GetCurrentMethod().Name.ToString(), "整理資料");

                    List<TxnTransactionMegre> mTransactionMegre =
                            (from p in MainServer.TxnAllData
                             group p by new { p.ISSUER_ID, p.CARD_TXN_TYPE_ID, p.CARD_TXN_SUBTYPE_ID, p.FARE_PRODUCT_TYPE_ID } into g
                             select new TxnTransactionMegre
                             {
                                 ISSUER_ID = g.Key.ISSUER_ID,
                                 CARD_TXN_TYPE_ID = g.Key.CARD_TXN_TYPE_ID,
                                 CARD_TXN_SUBTYPE_ID = g.Key.CARD_TXN_SUBTYPE_ID,
                                 FARE_PRODUCT_TYPE_ID = g.Key.FARE_PRODUCT_TYPE_ID,
                                 Count = g.Count(),
                                 Amt = g.Sum(p => Convert.ToInt16((p.TXN_AMT != "NULL") ? p.TXN_AMT : "0"))
                             }).ToList();

                    string msgTitle = $"" +
                        $"ISSUER_ID,IssureCodeName,CARD_TXN_TYPE_ID,TxnTypeName,CARD_TXN_SUBTYPE_ID,TxnTypeSubName,FARE_PRODUCT_TYPE_ID,FareProductTypeName,Count,Amt";
                    MyCommond.ExportData(TidyfileName, Encoding.UTF8, msgTitle); //個交易類型輸出的欄位名稱

                    //WriteLog(ExecutionMode.Simple, MethodBase.GetCurrentMethod().Name.ToString(), "分析資料");
                    int Pass1280Ride = 0;
                    int Pass1200Ride = 0;
                    int Pass1200Set = 0;
                    int Pass1200CancelCount = 0;
                    int Pass1200CancelAmt = 0;
                    int Pass1200RefundCount = 0;
                    int Pass1200RefundAmt = 0;

                    foreach (var item in mTransactionMegre)
                    {
                        #region 取得已經命名的ID名稱

                        //Console.WriteLine(item.ToString());
                        string IssureCodeName = "";
                        string TxnTypeName = "";
                        string TxnTypeSubName = "";
                        string FareProductTypeName = "";
                        bool TxnTypeCheck = (MainServer.globalCommond.Stock_Mrt_TxnType.Find(x => x.Num == Convert.ToInt16((item.CARD_TXN_TYPE_ID != "NULL") ? item.CARD_TXN_TYPE_ID : "0")).mCheck != "");
                        bool TxnTypeSubCheck = false;

                        if (MyCommond.ExportLanguage == Language.English)
                        {
                            IssureCodeName = (MainServer.globalCommond.Stock_Mrt_IssureCode.FindIndex(x => x.Num == Convert.ToInt16((item.ISSUER_ID != "NULL") ? item.ISSUER_ID : "0")) != -1) ? MainServer.globalCommond.Stock_Mrt_IssureCode.Find(x => x.Num == Convert.ToInt16(item.ISSUER_ID)).EngName : "";
                            TxnTypeName = (MainServer.globalCommond.Stock_Mrt_TxnType.FindIndex(x => x.Num == Convert.ToInt16((item.CARD_TXN_TYPE_ID != "NULL") ? item.CARD_TXN_TYPE_ID : "0")) != -1) ? MainServer.globalCommond.Stock_Mrt_TxnType.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_TYPE_ID)).EngName : "";
                            FareProductTypeName = (MainServer.globalCommond.Stock_Mrt_FareProductType.FindIndex(x => x.Num == Convert.ToInt16((item.FARE_PRODUCT_TYPE_ID != "NULL") ? item.FARE_PRODUCT_TYPE_ID : "0")) != -1) ? MainServer.globalCommond.Stock_Mrt_FareProductType.Find(x => x.Num == Convert.ToInt16(item.FARE_PRODUCT_TYPE_ID)).EngName : "";
                        }
                        else if (MyCommond.ExportLanguage == Language.zhTW)
                        {
                            IssureCodeName = (MainServer.globalCommond.Stock_Mrt_IssureCode.FindIndex(x => x.Num == Convert.ToInt16((item.ISSUER_ID != "NULL") ? item.ISSUER_ID : "0")) != -1) ? MainServer.globalCommond.Stock_Mrt_IssureCode.Find(x => x.Num == Convert.ToInt16(item.ISSUER_ID)).ChName : "";
                            TxnTypeName = (MainServer.globalCommond.Stock_Mrt_TxnType.FindIndex(x => x.Num == Convert.ToInt16((item.CARD_TXN_TYPE_ID != "NULL") ? item.CARD_TXN_TYPE_ID : "0")) != -1) ? MainServer.globalCommond.Stock_Mrt_TxnType.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_TYPE_ID)).ChName : "";
                            FareProductTypeName = (MainServer.globalCommond.Stock_Mrt_FareProductType.FindIndex(x => x.Num == Convert.ToInt16((item.FARE_PRODUCT_TYPE_ID != "NULL") ? item.FARE_PRODUCT_TYPE_ID : "0")) != -1) ? MainServer.globalCommond.Stock_Mrt_FareProductType.Find(x => x.Num == Convert.ToInt16(item.FARE_PRODUCT_TYPE_ID)).ChName : "";
                        }

                        if (MainServer.globalCommond.Stock_Mrt_TxnType.FindIndex(x => x.Num == Convert.ToInt16((item.CARD_TXN_TYPE_ID != "NULL") ? item.CARD_TXN_TYPE_ID : "0")) != -1)
                        {
                            if (MainServer.globalCommond.Stock_Mrt_IssureCode.Find(x => x.EngName == "ECC").Num == Convert.ToInt16((item.ISSUER_ID != "NULL") ? item.ISSUER_ID : "0"))
                            {
                                if (MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ecc.FindIndex(x => x.Num == Convert.ToInt16((item.CARD_TXN_SUBTYPE_ID != "NULL") ? item.CARD_TXN_SUBTYPE_ID : "0")) != -1)
                                {
                                    if (MyCommond.ExportLanguage == Language.English)
                                    {
                                        TxnTypeSubName = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ecc.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_SUBTYPE_ID)).EngName;
                                    }
                                    else if (MyCommond.ExportLanguage == Language.zhTW)
                                    {
                                        TxnTypeSubName = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ecc.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_SUBTYPE_ID)).ChName;
                                    }
                                    if (MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ecc.Find(x => x.Num == Convert.ToInt16((item.CARD_TXN_SUBTYPE_ID != "NULL") ? item.CARD_TXN_SUBTYPE_ID : "0")).mCheck != "")
                                    {
                                        TxnTypeSubCheck = true;
                                    }
                                }
                            }
                            else if (MainServer.globalCommond.Stock_Mrt_IssureCode.Find(x => x.EngName == "Ipass").Num == Convert.ToInt16((item.ISSUER_ID != "NULL") ? item.ISSUER_ID : "0"))
                            {
                                if (MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ipass.FindIndex(x => x.Num == Convert.ToInt16((item.CARD_TXN_SUBTYPE_ID != "NULL") ? item.CARD_TXN_SUBTYPE_ID : "0")) != -1)
                                {
                                    if (MyCommond.ExportLanguage == Language.English)
                                    {
                                        TxnTypeSubName = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ipass.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_SUBTYPE_ID)).EngName;
                                    }
                                    else if (MyCommond.ExportLanguage == Language.zhTW)
                                    {
                                        TxnTypeSubName = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ipass.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_SUBTYPE_ID)).ChName;
                                    }
                                    if (MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ipass.Find(x => x.Num == Convert.ToInt16((item.CARD_TXN_SUBTYPE_ID != "NULL") ? item.CARD_TXN_SUBTYPE_ID : "0")).mCheck != "")
                                    {
                                        TxnTypeSubCheck = true;
                                    }
                                }
                            }
                            else if (MainServer.globalCommond.Stock_Mrt_IssureCode.Find(x => x.EngName == "Icash").Num == Convert.ToInt16((item.ISSUER_ID != "NULL") ? item.ISSUER_ID : "0"))
                            {
                                if (MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Icash.FindIndex(x => x.Num == Convert.ToInt16((item.CARD_TXN_SUBTYPE_ID != "NULL") ? item.CARD_TXN_SUBTYPE_ID : "0")) != -1)
                                {
                                    if (MyCommond.ExportLanguage == Language.English)
                                    {
                                        TxnTypeSubName = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Icash.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_SUBTYPE_ID)).EngName;
                                    }
                                    else if (MyCommond.ExportLanguage == Language.zhTW)
                                    {
                                        TxnTypeSubName = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Icash.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_SUBTYPE_ID)).ChName;
                                    }
                                    if (MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Icash.Find(x => x.Num == Convert.ToInt16((item.CARD_TXN_SUBTYPE_ID != "NULL") ? item.CARD_TXN_SUBTYPE_ID : "0")).mCheck != "")
                                    {
                                        TxnTypeSubCheck = true;
                                    }
                                }
                            }
                            else
                            {
                                TxnTypeSubCheck = true;
                                if (MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ticket.FindIndex(x => x.Num == Convert.ToInt16((item.CARD_TXN_SUBTYPE_ID != "NULL") ? item.CARD_TXN_SUBTYPE_ID : "0")) != -1)
                                {
                                    if (MyCommond.ExportLanguage == Language.English)
                                    {
                                        TxnTypeSubName = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ticket.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_SUBTYPE_ID)).EngName;
                                    }
                                    else if (MyCommond.ExportLanguage == Language.zhTW)
                                    {
                                        TxnTypeSubName = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ticket.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_SUBTYPE_ID)).ChName;
                                    }
                                }
                            }
                        }
                        else
                        {

                        }

                        #endregion

                        #region 資料加總

                        if (TxnTypeCheck && TxnTypeSubCheck)
                        {
                            if (MainServer.globalCommond.Stock_Mrt_TxnType.Find(x => x.Num == Convert.ToInt16((item.CARD_TXN_TYPE_ID != "NULL") ? item.CARD_TXN_TYPE_ID : "0")).mCheck == "+")
                            {
                                RideCountAdd += item.Count;
                                AmtAdd += item.Amt;
                            }
                            else if (MainServer.globalCommond.Stock_Mrt_TxnType.Find(x => x.Num == Convert.ToInt16((item.CARD_TXN_TYPE_ID != "NULL") ? item.CARD_TXN_TYPE_ID : "0")).mCheck == "-")
                            {
                                RideCountMinus += item.Count;
                                AmtMinus += item.Amt;
                            }

                            if (Convert.ToInt16(item.CARD_TXN_TYPE_ID) == 4 && Convert.ToInt16(item.FARE_PRODUCT_TYPE_ID) == 74)
                            {
                                Pass1280Ride += item.Count;
                            }
                            else if (Convert.ToInt16(item.CARD_TXN_TYPE_ID) == 4 && Convert.ToInt16(item.FARE_PRODUCT_TYPE_ID) == 75)
                            {
                                Pass1200Ride += item.Count;
                            }
                        }

                        if (Convert.ToInt16(item.CARD_TXN_TYPE_ID) == 3 && Convert.ToInt16(item.FARE_PRODUCT_TYPE_ID) == 75)
                        {
                            Pass1200Set += item.Count;
                        }
                        else if (Convert.ToInt16(item.CARD_TXN_TYPE_ID) == 6 && Convert.ToInt16(item.FARE_PRODUCT_TYPE_ID) == 75)
                        {
                            Pass1200RefundCount += item.Count;
                            Pass1200RefundAmt += item.Amt;
                        }
                        else if (Convert.ToInt16(item.CARD_TXN_TYPE_ID) == 10 && Convert.ToInt16(item.FARE_PRODUCT_TYPE_ID) == 75)
                        {
                            Pass1200CancelCount += item.Count;
                            Pass1200CancelAmt += item.Amt;
                        }

                        #endregion

                        string msg1 = $"" +
                            $"{item.ISSUER_ID},{IssureCodeName},{item.CARD_TXN_TYPE_ID},{TxnTypeName},{item.CARD_TXN_SUBTYPE_ID},{TxnTypeSubName},{item.FARE_PRODUCT_TYPE_ID},{FareProductTypeName},{item.Count},{item.Amt}";
                        //WriteLog(ExecutionMode.Normal, MethodBase.GetCurrentMethod().Name.ToString(), msg1);
                        MyCommond.ExportData(TidyfileName, Encoding.UTF8, msg1);
                    }

                    //WriteLog(ExecutionMode.Simple, MethodBase.GetCurrentMethod().Name.ToString(), "輸出定期票");
                    string mPath_ReportFile = MyCommond.Path_Program + MyCommond.Path_Report + string.Format(MyCommond.Path_Date, MainServer.OperationDate.ToString("yyyy.MM"));
                    string ReportfileName = mPath_ReportFile + MyCommond.CheckAndReName(mPath_ReportFile, $"行政院通勤月票_環狀線_{MainServer.OperationDate:yyyyMMdd}", "csv");
                    string data = $"日期,路線,1280月票進站,1280月票出站,1200月票進站,1200月票出站,1200月票設定,1200月票取消,1200月票退票,1200月票退票金額,1200退票手續費,1200總設定筆數" + Environment.NewLine +
                        $"{MainServer.OperationDate:yyyyMMdd},{MainServer.globalCommond.OperationCode},N,{Pass1280Ride},N,{Pass1200Ride},{Pass1200Set},{Pass1200CancelCount},{Pass1200RefundCount},{Pass1200RefundAmt},N,{Pass1200Set - Pass1200CancelCount}";
                    //WriteLog(ExecutionMode.Simple, MethodBase.GetCurrentMethod().Name.ToString(), Environment.NewLine + data);
                    MyCommond.ExportData(ReportfileName, Encoding.UTF8, data);

                    DateTime EndTime = DateTime.Now;
                    CTime = EndTime - StartTime;
                    CountExit = RideCountAdd - RideCountMinus;
                    ExitAmt = AmtAdd - AmtMinus;

                    MessageBox.Show($"計算時間:{CTime}\n" +
                        $"運量 => 正數:{RideCountAdd,8}, 減數:{RideCountMinus,8}, 總計:{CountExit,8}\n" +
                        $"金額 => 正數:{AmtAdd,8}, 減數:{AmtMinus,8}, 總計:{ExitAmt,8}");

                    GC.Collect();
                }
                if (!true)
                {
                    MyCommond.WriteLog(ThisReceive, "查詢" + "所有交易類別分類總計");
                    //this.WriteLog(ExecutionMode.Normal, "查詢" + "所以有交易類別分類總計");

                    List<TxnTransactionMegre> StockMrtAllTxn = (
                        from p in MainServer.TxnAllData
                        where
                            (
                                MainServer.globalCommond.Stock_Mrt_Y_StationOd.FindIndex(x => x.ExitStation == p.SVCE_LOC_ID && x.IsCount == "Y") != -1
                                || mStation_0.Contains(Convert.ToInt16(p.SVCE_LOC_ID))
                            )
                            //MainServer.Stock_Mrt_StationOd.Find(x => 
                            //    x.EntryStation == MainServer.Stock_Mrt_Station.Find(y => y.CodeName == p.ENTRY_LOC_ID).StationName && 
                            //    x.ExitStation  == MainServer.Stock_Mrt_Station.Find(z => z.CodeName == p.SVCE_LOC_ID ).StationName
                            //).IsCount != ""
                            && Convert.ToDateTime(p.TXN_TIMESTAMP) >= MainServer.OperationDate
                            && Convert.ToDateTime(p.TXN_TIMESTAMP) < MainServer.OperationDate.AddDays(1)
                        group p by new 
                        {
                            p.ISSUER_ID, 
                            p.SVCE_LOC_ID, 
                            EQUIP_TYPE = p.DEV_ID.Substring(0, 3), 
                            EQUIP_ID = p.DEV_ID.Substring(3, 4), 
                            p.CARD_TXN_TYPE_ID, 
                            p.CARD_TXN_SUBTYPE_ID, 
                            p.FARE_PRODUCT_TYPE_ID, 
                            p.TICKET_SUBTYPE_ID, 
                            p.BUSINESS_DATE 
                        } into g
                        select new TxnTransactionMegre
                        {
                            ISSUER_ID = g.Key.ISSUER_ID,
                            SVCE_LOC_ID = g.Key.SVCE_LOC_ID,
                            EQUIP_TYPE = g.Key.EQUIP_TYPE,
                            EQUIP_ID = g.Key.EQUIP_ID,
                            CARD_TXN_TYPE_ID = g.Key.CARD_TXN_TYPE_ID,
                            CARD_TXN_SUBTYPE_ID = g.Key.CARD_TXN_SUBTYPE_ID,
                            FARE_PRODUCT_TYPE_ID = g.Key.FARE_PRODUCT_TYPE_ID,
                            TICKET_SUBTYPE_ID = g.Key.TICKET_SUBTYPE_ID,
                            BUSINESS_DATE = g.Key.BUSINESS_DATE,
                            Count = g.Count(),
                            Amt = g.Sum(p => Convert.ToInt16((p.TXN_AMT != "NULL") ? p.TXN_AMT : "0"))
                        }
                    ).ToList();

                    MyCommond.WriteLog(ThisReceive, "輸出查詢到的" + "所有交易類別分類總計");
                    MyCommond.ExportVerify<TxnTransactionMegre>(new ExportConfig()
                    {
                        Path1 = MyCommond.Path_Program + MyCommond.Path_Verify,
                        DTime = MainServer.OperationDate,
                        FileName = "StockMrtAllTxn",
                        LastName = "csv",
                        OpLine = MainServer.globalCommond.OperationCode_String,
                    }, ExportTxnTransactionMegreTitle, StockMrtAllTxn, true);

                    List<int> ls_O = new List<int>();
                    foreach (var item in MainServer.globalCommond.Stock_Mrt_TxnType)
                    {
                        if (item.mCheck != null && item.mCheck.Length == 2)
                        {
                            if (item.mCheck.Substring(1, 1) == "O")
                            {
                                ls_O.Add(item.Num);
                            }
                        }
                    }

                    List<int> ls_R = new List<int>();
                    foreach (var item in MainServer.globalCommond.Stock_Mrt_TxnType)
                    {
                        if (item.mCheck != null && item.mCheck.Length == 2)
                        {
                            if (item.mCheck.Substring(1, 1) == "R")
                            {
                                ls_R.Add(item.Num);
                            }
                        }
                    }

                    //WriteLog(ExecutionMode.Simple, MethodBase.GetCurrentMethod().Name.ToString(), "分析資料");
                    int Pass1280Ride = 0;
                    int Pass1200Ride = 0;
                    int Pass1200Set = 0;
                    int Pass1200CancelCount = 0;
                    int Pass1200CancelAmt = 0;
                    int Pass1200RefundCount = 0;
                    int Pass1200RefundAmt = 0;

                    foreach (var item in StockMrtAllTxn)
                    {
                        #region 取得已經命名的ID名稱

                        //Console.WriteLine(item.ToString());
                        string IssureCodeName = "";
                        string TxnTypeName = "";
                        string TxnTypeSubName = "";
                        string FareProductTypeName = "";
                        bool TxnTypeCheck = (MainServer.globalCommond.Stock_Mrt_TxnType.Find(x => x.Num == Convert.ToInt16((item.CARD_TXN_TYPE_ID != "NULL") ? item.CARD_TXN_TYPE_ID : "0")).mCheck != "");
                        bool TxnTypeSubCheck = false;

                        if (MyCommond.ExportLanguage == Language.English)
                        {
                            IssureCodeName = (MainServer.globalCommond.Stock_Mrt_IssureCode.FindIndex(x => x.Num == Convert.ToInt16((item.ISSUER_ID != "NULL") ? item.ISSUER_ID : "0")) != -1) ? MainServer.globalCommond.Stock_Mrt_IssureCode.Find(x => x.Num == Convert.ToInt16(item.ISSUER_ID)).EngName : "";
                            TxnTypeName = (MainServer.globalCommond.Stock_Mrt_TxnType.FindIndex(x => x.Num == Convert.ToInt16((item.CARD_TXN_TYPE_ID != "NULL") ? item.CARD_TXN_TYPE_ID : "0")) != -1) ? MainServer.globalCommond.Stock_Mrt_TxnType.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_TYPE_ID)).EngName : "";
                            FareProductTypeName = (MainServer.globalCommond.Stock_Mrt_FareProductType.FindIndex(x => x.Num == Convert.ToInt16((item.FARE_PRODUCT_TYPE_ID != "NULL") ? item.FARE_PRODUCT_TYPE_ID : "0")) != -1) ? MainServer.globalCommond.Stock_Mrt_FareProductType.Find(x => x.Num == Convert.ToInt16(item.FARE_PRODUCT_TYPE_ID)).EngName : "";
                        }
                        else if (MyCommond.ExportLanguage == Language.zhTW)
                        {
                            IssureCodeName = (MainServer.globalCommond.Stock_Mrt_IssureCode.FindIndex(x => x.Num == Convert.ToInt16((item.ISSUER_ID != "NULL") ? item.ISSUER_ID : "0")) != -1) ? MainServer.globalCommond.Stock_Mrt_IssureCode.Find(x => x.Num == Convert.ToInt16(item.ISSUER_ID)).ChName : "";
                            TxnTypeName = (MainServer.globalCommond.Stock_Mrt_TxnType.FindIndex(x => x.Num == Convert.ToInt16((item.CARD_TXN_TYPE_ID != "NULL") ? item.CARD_TXN_TYPE_ID : "0")) != -1) ? MainServer.globalCommond.Stock_Mrt_TxnType.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_TYPE_ID)).ChName : "";
                            FareProductTypeName = (MainServer.globalCommond.Stock_Mrt_FareProductType.FindIndex(x => x.Num == Convert.ToInt16((item.FARE_PRODUCT_TYPE_ID != "NULL") ? item.FARE_PRODUCT_TYPE_ID : "0")) != -1) ? MainServer.globalCommond.Stock_Mrt_FareProductType.Find(x => x.Num == Convert.ToInt16(item.FARE_PRODUCT_TYPE_ID)).ChName : "";
                        }

                        if (MainServer.globalCommond.Stock_Mrt_TxnType.FindIndex(x => x.Num == Convert.ToInt16((item.CARD_TXN_TYPE_ID != "NULL") ? item.CARD_TXN_TYPE_ID : "0")) != -1)
                        {
                            if (MainServer.globalCommond.Stock_Mrt_IssureCode.Find(x => x.EngName == "ECC").Num == Convert.ToInt16((item.ISSUER_ID != "NULL") ? item.ISSUER_ID : "0"))
                            {
                                if (MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ecc.FindIndex(x => x.Num == Convert.ToInt16((item.CARD_TXN_SUBTYPE_ID != "NULL") ? item.CARD_TXN_SUBTYPE_ID : "0")) != -1)
                                {
                                    if (MyCommond.ExportLanguage == Language.English)
                                    {
                                        TxnTypeSubName = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ecc.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_SUBTYPE_ID)).EngName;
                                    }
                                    else if (MyCommond.ExportLanguage == Language.zhTW)
                                    {
                                        TxnTypeSubName = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ecc.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_SUBTYPE_ID)).ChName;
                                    }
                                    if (MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ecc.Find(x => x.Num == Convert.ToInt16((item.CARD_TXN_SUBTYPE_ID != "NULL") ? item.CARD_TXN_SUBTYPE_ID : "0")).mCheck != "")
                                    {
                                        TxnTypeSubCheck = true;
                                    }
                                }
                            }
                            else if (MainServer.globalCommond.Stock_Mrt_IssureCode.Find(x => x.EngName == "Ipass").Num == Convert.ToInt16((item.ISSUER_ID != "NULL") ? item.ISSUER_ID : "0"))
                            {
                                if (MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ipass.FindIndex(x => x.Num == Convert.ToInt16((item.CARD_TXN_SUBTYPE_ID != "NULL") ? item.CARD_TXN_SUBTYPE_ID : "0")) != -1)
                                {
                                    if (MyCommond.ExportLanguage == Language.English)
                                    {
                                        TxnTypeSubName = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ipass.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_SUBTYPE_ID)).EngName;
                                    }
                                    else if (MyCommond.ExportLanguage == Language.zhTW)
                                    {
                                        TxnTypeSubName = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ipass.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_SUBTYPE_ID)).ChName;
                                    }
                                    if (MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ipass.Find(x => x.Num == Convert.ToInt16((item.CARD_TXN_SUBTYPE_ID != "NULL") ? item.CARD_TXN_SUBTYPE_ID : "0")).mCheck != "")
                                    {
                                        TxnTypeSubCheck = true;
                                    }
                                }
                            }
                            else if (MainServer.globalCommond.Stock_Mrt_IssureCode.Find(x => x.EngName == "Icash").Num == Convert.ToInt16((item.ISSUER_ID != "NULL") ? item.ISSUER_ID : "0"))
                            {
                                if (MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Icash.FindIndex(x => x.Num == Convert.ToInt16((item.CARD_TXN_SUBTYPE_ID != "NULL") ? item.CARD_TXN_SUBTYPE_ID : "0")) != -1)
                                {
                                    if (MyCommond.ExportLanguage == Language.English)
                                    {
                                        TxnTypeSubName = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Icash.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_SUBTYPE_ID)).EngName;
                                    }
                                    else if (MyCommond.ExportLanguage == Language.zhTW)
                                    {
                                        TxnTypeSubName = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Icash.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_SUBTYPE_ID)).ChName;
                                    }
                                    if (MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Icash.Find(x => x.Num == Convert.ToInt16((item.CARD_TXN_SUBTYPE_ID != "NULL") ? item.CARD_TXN_SUBTYPE_ID : "0")).mCheck != "")
                                    {
                                        TxnTypeSubCheck = true;
                                    }
                                }
                            }
                            else
                            {
                                TxnTypeSubCheck = true;
                                if (MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ticket.FindIndex(x => x.Num == Convert.ToInt16((item.CARD_TXN_SUBTYPE_ID != "NULL") ? item.CARD_TXN_SUBTYPE_ID : "0")) != -1)
                                {
                                    if (MyCommond.ExportLanguage == Language.English)
                                    {
                                        TxnTypeSubName = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ticket.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_SUBTYPE_ID)).EngName;
                                    }
                                    else if (MyCommond.ExportLanguage == Language.zhTW)
                                    {
                                        TxnTypeSubName = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ticket.Find(x => x.Num == Convert.ToInt16(item.CARD_TXN_SUBTYPE_ID)).ChName;
                                    }
                                }
                            }
                        }
                        else
                        {

                        }

                        #endregion

                        #region 資料加總

                        if (TxnTypeCheck && TxnTypeSubCheck)
                        {
                            if (true)
                            {
                                if (ls_O.Contains(Convert.ToInt16((item.CARD_TXN_TYPE_ID != "NULL") ? item.CARD_TXN_TYPE_ID : "0")))
                                { RideCountAdd += item.Count; AmtAdd += item.Amt; }
                                else if (ls_R.Contains(Convert.ToInt16((item.CARD_TXN_TYPE_ID != "NULL") ? item.CARD_TXN_TYPE_ID : "0")))
                                { RideCountMinus += item.Count; AmtMinus += item.Amt; }
                            }
                            else
                            {
                                if (MainServer.globalCommond.Stock_Mrt_TxnType.Find(x => x.Num == Convert.ToInt16((item.CARD_TXN_TYPE_ID != "NULL") ? item.CARD_TXN_TYPE_ID : "0")).mCheck == "+")
                                { RideCountAdd += item.Count; AmtAdd += item.Amt; }
                                else if (MainServer.globalCommond.Stock_Mrt_TxnType.Find(x => x.Num == Convert.ToInt16((item.CARD_TXN_TYPE_ID != "NULL") ? item.CARD_TXN_TYPE_ID : "0")).mCheck == "-")
                                { RideCountMinus += item.Count; AmtMinus += item.Amt; }
                            }

                            if (Convert.ToInt16(item.CARD_TXN_TYPE_ID) == 4 && Convert.ToInt16(item.FARE_PRODUCT_TYPE_ID) == 74)
                            { Pass1280Ride += item.Count; }
                            else if (Convert.ToInt16(item.CARD_TXN_TYPE_ID) == 4 && Convert.ToInt16(item.FARE_PRODUCT_TYPE_ID) == 75)
                            { Pass1200Ride += item.Count; }
                        }

                        if (Convert.ToInt16(item.CARD_TXN_TYPE_ID) == 3 && Convert.ToInt16(item.FARE_PRODUCT_TYPE_ID) == 75)
                        { Pass1200Set += item.Count; }
                        else if (Convert.ToInt16(item.CARD_TXN_TYPE_ID) == 6 && Convert.ToInt16(item.FARE_PRODUCT_TYPE_ID) == 75)
                        { Pass1200RefundCount += item.Count; Pass1200RefundAmt += item.Amt; }
                        else if (Convert.ToInt16(item.CARD_TXN_TYPE_ID) == 10 && Convert.ToInt16(item.FARE_PRODUCT_TYPE_ID) == 75)
                        { Pass1200CancelCount += item.Count; Pass1200CancelAmt += item.Amt; }

                        #endregion

                        string msg1 = $"" +
                            $"{item.ISSUER_ID},{IssureCodeName},{item.CARD_TXN_TYPE_ID},{TxnTypeName},{item.CARD_TXN_SUBTYPE_ID},{TxnTypeSubName},{item.FARE_PRODUCT_TYPE_ID},{FareProductTypeName},{item.Count},{item.Amt}";
                        //WriteLog(ExecutionMode.Normal, MethodBase.GetCurrentMethod().Name.ToString(), msg1);
                        MyCommond.ExportData(TidyfileName, Encoding.UTF8, msg1);
                    }

                    //WriteLog(ExecutionMode.Simple, MethodBase.GetCurrentMethod().Name.ToString(), "輸出定期票");
                    string mPath_ReportFile = MyCommond.Path_Program + MyCommond.Path_Report + string.Format(MyCommond.Path_Date, MainServer.OperationDate.ToString("yyyy.MM"));
                    string ReportfileName = mPath_ReportFile + MyCommond.CheckAndReName(mPath_ReportFile, $"行政院通勤月票_環狀線_{MainServer.OperationDate:yyyyMMdd}", "csv");
                    string data = $"日期,路線,1280月票進站,1280月票出站,1200月票進站,1200月票出站,1200月票設定,1200月票取消,1200月票退票,1200月票退票金額,1200退票手續費,1200總設定筆數" + Environment.NewLine +
                        $"{MainServer.OperationDate:yyyyMMdd},{MainServer.globalCommond.OperationCode},N,{Pass1280Ride},N,{Pass1200Ride},{Pass1200Set},{Pass1200CancelCount},{Pass1200RefundCount},{Pass1200RefundAmt},N,{Pass1200Set - Pass1200CancelCount}";
                    //WriteLog(ExecutionMode.Simple, MethodBase.GetCurrentMethod().Name.ToString(), Environment.NewLine + data);
                    MyCommond.ExportData(ReportfileName, Encoding.UTF8, data);

                    DateTime EndTime = DateTime.Now;
                    CTime = EndTime - StartTime;
                    CountExit = RideCountAdd - RideCountMinus;
                    ExitAmt = AmtAdd - AmtMinus;

                    MessageBox.Show($"計算時間:{CTime}\n" +
                        $"運量 => 正數:{RideCountAdd,8}, 減數:{RideCountMinus,8}, 總計:{CountExit,8}\n" +
                        $"金額 => 正數:{AmtAdd,8}, 減數:{AmtMinus,8}, 總計:{ExitAmt,8}");

                    StockMrtAllTxn = null;
                } // 各資料類型

                    #endregion
            }
            catch (Exception ex)
            {
                throw ex;
            }
            MyCommond.WriteLog(ThisReceive, $"簡易資料計算 - 結束");
            Status = StatusEnum.Finish;
        }

        #endregion

        //private string RowDataFolder = @"RowData\";
        private string RowDataFolder = @"MRT\";

        #region 各交易資料

        private void Calculator_TxnTypeAnalyze()
        {
            Status = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"各交易資料 - 開始");
            string ExportTxnTransactionTitle = (new TxnTransaction()).ExportTitle_English();

            CardSN = MainServer.OperationDate.ToString("yyyyMMdd");
            string ExportPath = MyCommond.Path_Program + MyCommond.Path_Report + RowDataFolder;
            string ExpofrLastFileName = "csv";
            string ExportOp = MainServer.globalCommond.OperationCode_String;
            DateTime DT = MainServer.OperationDate;
            string FName = "";

            try
            {
                #region 加值
                MyCommond.WriteLog(ThisReceive, "查詢" + "加值");
                FName = "AddValue";
                //this.WriteLog(ExecutionMode.Normal, "查詢" + "加值");

                List<TxnTransaction> StockMrtAddValue = (
                    from p in MainServer.TxnAllData
                    where
                        (
                            MainServer.globalCommond.Stock_Mrt_Y_StationOd.FindIndex(x => x.ExitStation == p.SVCE_LOC_ID && x.IsCount == "Y") != -1
                            || mStation_0.Contains(Convert.ToInt16(p.SVCE_LOC_ID))
                        )
                        && Convert.ToDateTime(p.TXN_TIMESTAMP) >= MainServer.OperationDate
                        && Convert.ToDateTime(p.TXN_TIMESTAMP) < MainServer.OperationDate.AddDays(1)
                        && MainServer.globalCommond.Stock_Mrt_TxnType.FindIndex(x => Convert.ToString(x.Num) == p.CARD_TXN_TYPE_ID && x.mCheck == "AA") != -1
                    orderby p.TXN_TIMESTAMP
                    select p
                ).ToList();

                MyCommond.WriteLog(ThisReceive, "查詢" + "加值");

                MyCommond.ExportData<TxnTransaction>(ThisReceive, new ExportConfig()
                {
                    Path1 = ExportPath + FName + @"\",
                    DTime = DT,
                    FirstName = "",
                    FileName = FName,
                    SubName = "",
                    OpLine = ExportOp,
                    LastName = ExpofrLastFileName,
                }, new TxnTransaction().ExportTitle_zhTW(), StockMrtAddValue, Encoding.UTF8);

                StockMrtAddValue = null;

                #endregion
                #region 取消加值
                MyCommond.WriteLog(ThisReceive, "查詢" + "取消加值");
                FName = "AddValueCancel";
                //this.WriteLog(ExecutionMode.Normal, "查詢" + "取消加值");

                List<TxnTransaction> StockMrtAddValueCancel = (
                    from p in MainServer.TxnAllData
                    where
                        (
                            MainServer.globalCommond.Stock_Mrt_Y_StationOd.FindIndex(x => x.ExitStation == p.SVCE_LOC_ID && x.IsCount == "Y") != -1
                            || mStation_0.Contains(Convert.ToInt16(p.SVCE_LOC_ID))
                        )
                        && Convert.ToDateTime(p.TXN_TIMESTAMP) >= MainServer.OperationDate
                        && Convert.ToDateTime(p.TXN_TIMESTAMP) < MainServer.OperationDate.AddDays(1)
                        && MainServer.globalCommond.Stock_Mrt_TxnType.FindIndex(x => Convert.ToString(x.Num) == p.CARD_TXN_TYPE_ID && x.mCheck == "AF") != -1
                    orderby p.TXN_TIMESTAMP
                    select p
                ).ToList();

                MyCommond.WriteLog(ThisReceive, "查詢" + "取消加值");

                MyCommond.ExportData<TxnTransaction>(ThisReceive, new ExportConfig()
                {
                    Path1 = ExportPath + FName + @"\",
                    DTime = DT,
                    FirstName = "",
                    FileName = FName,
                    SubName = "",
                    OpLine = ExportOp,
                    LastName = ExpofrLastFileName,
                }, new TxnTransaction().ExportTitle_zhTW(), StockMrtAddValueCancel, Encoding.UTF8);

                StockMrtAddValueCancel = null;

                #endregion
                #region 售卡
                MyCommond.WriteLog(ThisReceive, "查詢" + "售卡");
                FName = "SaleCard";
                //this.WriteLog(ExecutionMode.Normal, "查詢" + "售卡");

                List<TxnTransaction> StockMrtSaleCard = (
                    from p in MainServer.TxnAllData
                    where
                        (
                            MainServer.globalCommond.Stock_Mrt_Y_StationOd.FindIndex(x => x.ExitStation == p.SVCE_LOC_ID && x.IsCount == "Y") != -1
                            || mStation_0.Contains(Convert.ToInt16(p.SVCE_LOC_ID))
                        )
                        && Convert.ToDateTime(p.TXN_TIMESTAMP) >= MainServer.OperationDate
                        && Convert.ToDateTime(p.TXN_TIMESTAMP) < MainServer.OperationDate.AddDays(1)
                        && MainServer.globalCommond.Stock_Mrt_TxnType.FindIndex(x => Convert.ToString(x.Num) == p.CARD_TXN_TYPE_ID && x.mCheck == "CA") != -1
                    orderby p.TXN_TIMESTAMP
                    select p
                ).ToList();

                MyCommond.WriteLog(ThisReceive, "查詢" + "售卡");

                MyCommond.ExportData<TxnTransaction>(ThisReceive, new ExportConfig()
                {
                    Path1 = ExportPath + FName + @"\",
                    DTime = DT,
                    FirstName = "",
                    FileName = FName,
                    SubName = "",
                    OpLine = ExportOp,
                    LastName = ExpofrLastFileName,
                }, new TxnTransaction().ExportTitle_zhTW(), StockMrtSaleCard, Encoding.UTF8);

                StockMrtSaleCard = null;

                #endregion
                #region 取消售卡
                MyCommond.WriteLog(ThisReceive, "查詢" + "取消售卡");
                FName = "SaleCardCancel";
                //this.WriteLog(ExecutionMode.Normal, "查詢" + "取消售卡");

                List<TxnTransaction> StockMrtCardCancel = (
                        from p in MainServer.TxnAllData
                        where
                            (
                                MainServer.globalCommond.Stock_Mrt_Y_StationOd.FindIndex(x => x.ExitStation == p.SVCE_LOC_ID && x.IsCount == "Y") != -1
                                || mStation_0.Contains(Convert.ToInt16(p.SVCE_LOC_ID))
                            )
                            && Convert.ToDateTime(p.TXN_TIMESTAMP) >= MainServer.OperationDate
                            && Convert.ToDateTime(p.TXN_TIMESTAMP) < MainServer.OperationDate.AddDays(1)
                            && MainServer.globalCommond.Stock_Mrt_TxnType.FindIndex(x => Convert.ToString(x.Num) == p.CARD_TXN_TYPE_ID && x.mCheck == "CF") != -1
                        orderby p.TXN_TIMESTAMP
                        select p
                ).ToList();

                MyCommond.WriteLog(ThisReceive, "查詢" + "取消售卡");

                MyCommond.ExportData<TxnTransaction>(ThisReceive, new ExportConfig()
                {
                    Path1 = ExportPath + FName + @"\",
                    DTime = DT,
                    FirstName = "",
                    FileName = FName,
                    SubName = "",
                    OpLine = ExportOp,
                    LastName = ExpofrLastFileName,
                }, new TxnTransaction().ExportTitle_zhTW(), StockMrtCardCancel, Encoding.UTF8);

                StockMrtCardCancel = null;

                #endregion
                #region 售票
                MyCommond.WriteLog(ThisReceive, "查詢" + "售票");
                FName = "SaleTicket";
                //this.WriteLog(ExecutionMode.Normal, "查詢" + "售票");

                List<TxnTransaction> StockMrtSaleTicket = (
                        from p in MainServer.TxnAllData
                        where
                            (
                                MainServer.globalCommond.Stock_Mrt_Y_StationOd.FindIndex(x => x.ExitStation == p.SVCE_LOC_ID && x.IsCount == "Y") != -1
                                || mStation_0.Contains(Convert.ToInt16(p.SVCE_LOC_ID))
                            )
                            && Convert.ToDateTime(p.TXN_TIMESTAMP) >= MainServer.OperationDate
                            && Convert.ToDateTime(p.TXN_TIMESTAMP) < MainServer.OperationDate.AddDays(1)
                            && MainServer.globalCommond.Stock_Mrt_TxnType.FindIndex(x => Convert.ToString(x.Num) == p.CARD_TXN_TYPE_ID && x.mCheck == "TA") != -1
                        orderby p.TXN_TIMESTAMP
                        select p
                ).ToList();

                MyCommond.WriteLog(ThisReceive, "查詢" + "售票");

                MyCommond.ExportData<TxnTransaction>(ThisReceive, new ExportConfig()
                {
                    Path1 = ExportPath + FName + @"\",
                    DTime = DT,
                    FirstName = "",
                    FileName = FName,
                    SubName = "",
                    OpLine = ExportOp,
                    LastName = ExpofrLastFileName,
                }, new TxnTransaction().ExportTitle_zhTW(), StockMrtSaleTicket, Encoding.UTF8);

                StockMrtSaleTicket = null;

                #endregion
                #region 取消售票
                MyCommond.WriteLog(ThisReceive, "查詢" + "取消售票");
                FName = "SaleTicketCancel";
                //this.WriteLog(ExecutionMode.Normal, "查詢" + "取消售票");

                List<TxnTransaction> StockMrtTicketCancel = (
                        from p in MainServer.TxnAllData
                        where
                            (
                                MainServer.globalCommond.Stock_Mrt_Y_StationOd.FindIndex(x => x.ExitStation == p.SVCE_LOC_ID && x.IsCount == "Y") != -1
                                || mStation_0.Contains(Convert.ToInt16(p.SVCE_LOC_ID))
                            )
                            && Convert.ToDateTime(p.TXN_TIMESTAMP) >= MainServer.OperationDate
                            && Convert.ToDateTime(p.TXN_TIMESTAMP) < MainServer.OperationDate.AddDays(1)
                            && MainServer.globalCommond.Stock_Mrt_TxnType.FindIndex(x => Convert.ToString(x.Num) == p.CARD_TXN_TYPE_ID && x.mCheck == "TF") != -1
                        orderby p.TXN_TIMESTAMP
                        select p
                ).ToList();

                MyCommond.WriteLog(ThisReceive, "輸出" + "取消售票");

                MyCommond.ExportData<TxnTransaction>(ThisReceive, new ExportConfig()
                {
                    Path1 = ExportPath + FName + @"\",
                    DTime = DT,
                    FirstName = "",
                    FileName = FName,
                    SubName = "",
                    OpLine = ExportOp,
                    LastName = ExpofrLastFileName,
                }, new TxnTransaction().ExportTitle_zhTW(), StockMrtTicketCancel, Encoding.UTF8);

                StockMrtTicketCancel = null;

                #endregion
                #region 退票
                MyCommond.WriteLog(ThisReceive, "查詢" + "退票");
                FName = "SaleTicketRefound";
                //this.WriteLog(ExecutionMode.Normal, "查詢" + "退票");

                List<TxnTransaction> StockMrtTicketRefound = (
                    from p in MainServer.TxnAllData
                    where
                        (
                            MainServer.globalCommond.Stock_Mrt_Y_StationOd.FindIndex(x => x.ExitStation == p.SVCE_LOC_ID && x.IsCount == "Y") != -1
                            || mStation_0.Contains(Convert.ToInt16(p.SVCE_LOC_ID))
                        )
                        && Convert.ToDateTime(p.TXN_TIMESTAMP) >= MainServer.OperationDate
                        && Convert.ToDateTime(p.TXN_TIMESTAMP) < MainServer.OperationDate.AddDays(1)
                        && MainServer.globalCommond.Stock_Mrt_TxnType.FindIndex(x => Convert.ToString(x.Num) == p.CARD_TXN_TYPE_ID && x.mCheck == "TR") != -1
                    orderby p.TXN_TIMESTAMP
                    select p
                ).ToList();

                MyCommond.WriteLog(ThisReceive, "輸出" + "退票");

                MyCommond.ExportData<TxnTransaction>(ThisReceive, new ExportConfig()
                {
                    Path1 = ExportPath + FName + @"\",
                    DTime = DT,
                    FirstName = "",
                    FileName = FName,
                    SubName = "",
                    OpLine = ExportOp,
                    LastName = ExpofrLastFileName,
                }, new TxnTransaction().ExportTitle_zhTW(), StockMrtTicketRefound, Encoding.UTF8);

                StockMrtTicketRefound = null;

                #endregion
                #region 各交易計數
                MyCommond.WriteLog(ThisReceive, "查詢" + "所有交易類別合併總計");
                FName = "StockMrtAllTxn";
                List<TxnTransactionMegre> StockMrtAllTxn = (
                    from p in MainServer.TxnAllData
                    where
                        (
                            MainServer.globalCommond.Stock_Mrt_Y_StationOd.FindIndex(x => x.ExitStation == p.SVCE_LOC_ID && x.IsCount == "Y") != -1
                            || mStation_0.Contains(Convert.ToInt16(p.SVCE_LOC_ID))
                        )
                    group p by new
                    {
                        p.ISSUER_ID,
                        p.SVCE_LOC_ID,
                        EQUIP_TYPE = p.DEV_ID.Substring(0, 3),
                        EQUIP_ID = p.DEV_ID.Substring(3, 4),
                        p.CARD_TXN_TYPE_ID,
                        p.CARD_TXN_SUBTYPE_ID,
                        p.FARE_PRODUCT_TYPE_ID,
                        p.TICKET_SUBTYPE_ID,
                        p.BUSINESS_DATE
                    } into g
                    select new TxnTransactionMegre
                    {
                        ISSUER_ID = g.Key.ISSUER_ID,
                        SVCE_LOC_ID = g.Key.SVCE_LOC_ID,
                        EQUIP_TYPE = g.Key.EQUIP_TYPE,
                        EQUIP_ID = g.Key.EQUIP_ID,
                        CARD_TXN_TYPE_ID = g.Key.CARD_TXN_TYPE_ID,
                        CARD_TXN_SUBTYPE_ID = g.Key.CARD_TXN_SUBTYPE_ID,
                        FARE_PRODUCT_TYPE_ID = g.Key.FARE_PRODUCT_TYPE_ID,
                        TICKET_SUBTYPE_ID = g.Key.TICKET_SUBTYPE_ID,
                        BUSINESS_DATE = g.Key.BUSINESS_DATE,
                        Count = g.Count(),
                        Amt = g.Sum(p => Convert.ToInt16((p.TXN_AMT != "NULL") ? p.TXN_AMT : "0"))
                    }
                ).ToList();

                MyCommond.ExportData<TxnTransactionMegre>(ThisReceive, new ExportConfig()
                {
                    Path1 = ExportPath + FName + @"\",
                    DTime = DT,
                    FirstName = "",
                    FileName = FName,
                    SubName = "",
                    OpLine = ExportOp,
                    LastName = ExpofrLastFileName,
                }, new TxnTransactionMegre().ExportTitle_zhTW(), StockMrtAllTxn, Encoding.UTF8);
                #endregion
            }
            catch (Exception ex)
            {
                throw ex;
            }
            MyCommond.WriteLog(ThisReceive, $"各交易資料 - 結束");
            Status = StatusEnum.Finish;
        }

        #endregion

        #region 進出站

        private void Calculator_IOAnalyze()
        {
            Status = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"各進出站分類 - 開始");
            string ExportTxnTransactionTitle = (new TxnTransaction()).ExportTitle_zhTW();

            CardSN = MainServer.OperationDate.ToString("yyyyMMdd");
            string ExportPath = MyCommond.Path_Program + MyCommond.Path_Report + RowDataFolder;
            DateTime DT = MainServer.OperationDate;
            string FName = "";
            string ExportOp = MainServer.globalCommond.OperationCode_String;
            string ExpofrLastFileName = "csv";
            try
            {
                #region 新北進 新北出
                MyCommond.WriteLog(ThisReceive, "查詢" + "新北進 新北出");
                FName = "新北進新北出";
                List<TxnTransaction> Ntmc_I_Ntmc_O = (
                    from p in MainServer.TxnAllData
                    where
                        mTxnType.Contains(MainServer.globalCommond.Stock_Mrt_TxnType.Find(x => x.Num == Convert.ToInt16(p.CARD_TXN_TYPE_ID)).mCheck) &&
                        ((p.ENTRY_LOC_ID != "NULL") ? Convert.ToInt16(p.ENTRY_LOC_ID) >= 200 : false) &&
                        Convert.ToInt16(p.SVCE_LOC_ID) >= 200
                    orderby p.TXN_TIMESTAMP
                    select p
                ).ToList();

                MyCommond.WriteLog(ThisReceive, "輸出查詢到的" + "新北進 新北出");

                MyCommond.ExportData<TxnTransaction>(ThisReceive, new ExportConfig()
                {
                    Path1 = ExportPath + FName + @"\",
                    DTime = DT,
                    FirstName = "",
                    FileName = FName,
                    SubName = "",
                    OpLine = ExportOp,
                    LastName = ExpofrLastFileName,
                }, ExportTxnTransactionTitle, Ntmc_I_Ntmc_O, Encoding.UTF8);

                Ntmc_I_Ntmc_O = null;
                #endregion
                #region 新北進 北捷出
                MyCommond.WriteLog(ThisReceive, "查詢" + "新北進 北捷出");
                FName = "新北進北捷出";
                List<TxnTransaction> Metro_I_Ntmc_O = (
                    from p in MainServer.TxnAllData
                    where
                        mTxnType.Contains(MainServer.globalCommond.Stock_Mrt_TxnType.Find(x => x.Num == Convert.ToInt16(p.CARD_TXN_TYPE_ID)).mCheck) &&
                        ((p.ENTRY_LOC_ID != "NULL") ? mStation_1.Contains(Convert.ToInt16(p.ENTRY_LOC_ID)) : false) &&
                        Convert.ToInt16(p.SVCE_LOC_ID) < 200
                    orderby p.TXN_TIMESTAMP
                    select p
                ).ToList();

                MyCommond.WriteLog(ThisReceive, "輸出查詢到的" + "新北進 北捷出");

                MyCommond.ExportData<TxnTransaction>(ThisReceive, new ExportConfig()
                {
                    Path1 = ExportPath + FName + @"\",
                    DTime = DT,
                    FirstName = "",
                    FileName = FName,
                    SubName = "",
                    OpLine = ExportOp,
                    LastName = ExpofrLastFileName,
                }, ExportTxnTransactionTitle, Metro_I_Ntmc_O, Encoding.UTF8);

                Metro_I_Ntmc_O = null;
                #endregion
                #region 北捷進 新北出
                MyCommond.WriteLog(ThisReceive, "查詢" + "北捷進 新北出");
                FName = "北捷進新北出";
                List<TxnTransaction> Ntmc_I_Metro_O = (
                    from p in MainServer.TxnAllData
                    where
                        mTxnType.Contains(MainServer.globalCommond.Stock_Mrt_TxnType.Find(x => x.Num == Convert.ToInt16(p.CARD_TXN_TYPE_ID)).mCheck) &&
                        ((p.ENTRY_LOC_ID != "NULL") ? Convert.ToInt16(p.ENTRY_LOC_ID) < 200 : false) &&
                        mStation_1.Contains(Convert.ToInt16(p.SVCE_LOC_ID))
                    orderby p.TXN_TIMESTAMP
                    select p
                ).ToList();

                MyCommond.WriteLog(ThisReceive, "輸出查詢到的" + "北捷進 新北出");
            
                MyCommond.ExportData<TxnTransaction>(ThisReceive, new ExportConfig()
                {
                    Path1 = ExportPath + FName + @"\",
                    DTime = DT,
                    FirstName = "",
                    FileName = FName,
                    SubName = "",
                    OpLine = ExportOp,
                    LastName = ExpofrLastFileName,
                }, ExportTxnTransactionTitle, Ntmc_I_Metro_O, Encoding.UTF8);

                Ntmc_I_Metro_O = null;
                #endregion
                #region 北捷進 北捷出
                MyCommond.WriteLog(ThisReceive, "查詢" + "北捷進 北捷出");
                FName = "北捷進北捷出";
                List<TxnTransaction> Metro_I_Metro_O = (
                    from p in MainServer.TxnAllData
                    where
                        mTxnType.Contains(MainServer.globalCommond.Stock_Mrt_TxnType.Find(x => x.Num == Convert.ToInt16(p.CARD_TXN_TYPE_ID)).mCheck) &&
                        ((p.ENTRY_LOC_ID != "NULL") ? Convert.ToInt16(p.ENTRY_LOC_ID) < 200 : false) &&
                        Convert.ToInt16(p.SVCE_LOC_ID) < 200
                    orderby p.TXN_TIMESTAMP
                    select p
                ).ToList();

                MyCommond.WriteLog(ThisReceive, "輸出查詢到的" + "北捷進 北捷出");
            
                MyCommond.ExportData<TxnTransaction>(ThisReceive, new ExportConfig()
                {
                    Path1 = ExportPath + FName + @"\",
                    DTime = DT,
                    FirstName = "",
                    FileName = FName,
                    SubName = "",
                    OpLine = ExportOp,
                    LastName = ExpofrLastFileName,
                }, ExportTxnTransactionTitle, Metro_I_Metro_O, Encoding.UTF8);

                Metro_I_Metro_O = null;
                #endregion
                #region 進站NULL 新北出
                MyCommond.WriteLog(ThisReceive, "查詢" + "進站NULL 新北出");
                FName = "進站NULL新北出";
                List<TxnTransaction> Null_I_Ntmc_O = (
                    from p in MainServer.TxnAllData
                    where
                        mTxnType.Contains(MainServer.globalCommond.Stock_Mrt_TxnType.Find(x => x.Num == Convert.ToInt16(p.CARD_TXN_TYPE_ID)).mCheck) &&
                        p.ENTRY_LOC_ID == "NULL" &&
                        Convert.ToInt16(p.SVCE_LOC_ID) >= 200
                    orderby p.TXN_TIMESTAMP
                    select p
                ).ToList();

                MyCommond.WriteLog(ThisReceive, "輸出查詢到的" + "進站NULL 新北出");
            
                MyCommond.ExportData<TxnTransaction>(ThisReceive, new ExportConfig()
                {
                    Path1 = ExportPath + FName + @"\",
                    DTime = DT,
                    FirstName = "",
                    FileName = FName,
                    SubName = "",
                    OpLine = ExportOp,
                    LastName = ExpofrLastFileName,
                }, ExportTxnTransactionTitle, Null_I_Ntmc_O, Encoding.UTF8);

                Null_I_Ntmc_O = null;
                #endregion
                #region 進站NULL 北捷出
                MyCommond.WriteLog(ThisReceive, "查詢" + "進站NULL 北捷出");
                FName = "進站NULL北捷出";
                List<TxnTransaction> Null_I_Metro_O = (
                    from p in MainServer.TxnAllData
                    where
                        mTxnType.Contains(MainServer.globalCommond.Stock_Mrt_TxnType.Find(x => x.Num == Convert.ToInt16(p.CARD_TXN_TYPE_ID)).mCheck) &&
                        p.ENTRY_LOC_ID == "NULL" &&
                        Convert.ToInt16(p.SVCE_LOC_ID) < 200
                    orderby p.TXN_TIMESTAMP
                    select p
                ).ToList();

                MyCommond.WriteLog(ThisReceive, "輸出查詢到的" + "進站NULL 北捷出");
            
                MyCommond.ExportData<TxnTransaction>(ThisReceive, new ExportConfig()
                {
                    Path1 = ExportPath + FName + @"\",
                    DTime = DT,
                    FirstName = "",
                    FileName = FName,
                    SubName = "",
                    OpLine = ExportOp,
                    LastName = ExpofrLastFileName,
                }, ExportTxnTransactionTitle, Null_I_Metro_O, Encoding.UTF8);
                Null_I_Metro_O = null;
                    #endregion
            }
            catch (Exception ex)
            {
                throw ex;
            }

            MyCommond.WriteLog(ThisReceive, $"各進出站分類 - 結束");
            Status = StatusEnum.Finish;
        }

        #endregion

        #region 合併里程(延人公里)

        private void Calculator_CompareAnalyze_Old()
        {
            Console.WriteLine("New Card Txn");
            Status = StatusEnum.Calculating;
            CardSN = MainServer.OperationDate.ToString("yyyyMMdd");
            TxnTransaction StockTxn_1 = null;
            TxnTransaction StockTxn_2 = null;
            TxnTransaction StockTxn_3 = new TxnTransaction();
            try
            {
            ColculatorStart1:
                List<TxnTransaction> StockExportTxn_0 = null;
                Status = StatusEnum.Calculating;

                try
                {
                    StockExportTxn_0 = CardTxnData.ToList(); /* 存放移除過後的資料  (Error => StockExportTxn_0的資料被移除時StockData資料也被移除) */
                }
                catch (Exception ex)
                {
                    Status = StatusEnum.ReGetCardTxn;
                    GetCardTxn();
                    while (CardTxnData.Count() < 1) { Thread.Sleep(30); }
                    goto ColculatorStart1;
                }

                MyCommond.WriteLog(ThisReceive, $"判讀第{Counts}張({CardSN,20})開始");
                List<TxnTransaction> StockExportTxn_1 = new List<TxnTransaction>();  // 存放移除的資料
                List<TxnTransaction> StockExportTxn_2 = new List<TxnTransaction>();  // 存放合併的資料
                List<TxnTransaction> StockExportTxn_3 = new List<TxnTransaction>();  // 存放異常資料

                string[] TxnType_In = { "20", "151", "153", "191", "194" };
                string[] TxnType_Out = { "1", "2", "4", "8", "13", "152", "154", "192", "195" };
                List<MrtStationList_Old> Stock_Mrt_Station = new List<MrtStationList_Old>();
                TimeSpan OverTime;
                
                //因2024/04/03大地震 影響板橋至中和段段軌，以接駁車載運，計算用
                OverTime = new TimeSpan(1, 0, 0);
                Stock_Mrt_Station.Add(new MrtStationList_Old() { CodeName = "82", StationName = "板橋", CodeName2 = "209", StationName2 = "板橋" });
                Stock_Mrt_Station.Add(new MrtStationList_Old() { CodeName = "205", StationName = "中和", CodeName2 = "205", StationName2 = "中和" });

                /// //原延人里程用
                /// Stock_Mrt_Station.Add(new MrtStationList_Old() { CodeName = "82", StationName = "板橋", CodeName2 = "209", StationName2 = "板橋" });
                /// Stock_Mrt_Station.Add(new MrtStationList_Old() { CodeName = "83", StationName = "新埔", CodeName2 = "210", StationName2 = "新埔民生" });
                /// Stock_Mrt_Station.Add(new MrtStationList_Old() { CodeName = "209", StationName = "板橋", CodeName2 = "82", StationName2 = "板橋" });
                /// Stock_Mrt_Station.Add(new MrtStationList_Old() { CodeName = "210", StationName = "新埔民生", CodeName2 = "83", StationName2 = "新埔" });
                /// 
                /// OverTime = new TimeSpan(0, 30, 0);

                int StepCount = 0;

                foreach (var item in CardTxnData)
                {
                    if (StockExportTxn_0.Count() < 3) continue;
                    Console.WriteLine($"{StepCount++,4}, {item.ToString()}");
                    bool FindIn = TxnType_In.Contains(item.CARD_TXN_TYPE_ID);
                    bool FindOut = TxnType_Out.Contains(item.CARD_TXN_TYPE_ID);

                    Console.WriteLine($"FindIn {FindIn}, FindOut {FindOut}");

                    if (FindIn)
                    {
                        Console.WriteLine("In");
                        if (StockTxn_1 == null)
                        {
                            Console.WriteLine("In_1");
                            continue;
                        }
                        else if (StockTxn_2 == null)
                        {
                            bool B_Time = (Convert.ToDateTime(StockTxn_1.TXN_TIMESTAMP) - Convert.ToDateTime(item.TXN_TIMESTAMP)) < OverTime;
                            bool B_Station = MainServer.globalCommond.Stock_Mrt_Y_Station_Old.Find(x => x.CodeName == StockTxn_1.SVCE_LOC_ID && x.CodeName2 == item.SVCE_LOC_ID) != null;

                            if (B_Time && B_Station) { Console.WriteLine("In_2-1"); StockTxn_2 = item; }
                            else { Console.WriteLine("In_2-2"); StockTxn_1 = null; }
                            continue;
                        }
                        else { StockTxn_1 = null; StockTxn_2 = null; continue; }
                    }
                    else if (FindOut)
                    {
                        Console.WriteLine("Out");
                        if (StockTxn_1 == null)
                        {
                            Console.WriteLine("Out_1");
                            StockTxn_1 = item;
                            continue;
                        }
                        else if (StockTxn_2 == null)
                        {
                            Console.WriteLine("Out_2");
                            StockTxn_1 = item;
                            continue;
                        }
                        else
                        {
                            Console.WriteLine("Out_3");

                            //將已經輸出過個資料移除
                            StockExportTxn_0.Remove(StockTxn_1);
                            StockExportTxn_0.Remove(StockTxn_2);
                            StockExportTxn_0.Remove(item);

                            if (StockTxn_1.ENTRY_LOC_ID != "NULL")
                            {
                                ExportTxnCompare(StockTxn_1);
                                ExportTxnCompare(StockTxn_2);
                                ExportTxnCompare(item);

                                StockTxn_3 = item;
                                StockTxn_3.TXN_AMT = Convert.ToString(Convert.ToInt16(StockTxn_1.TXN_AMT) + Convert.ToInt16(item.TXN_AMT));  //之後看到定期票要重算金額
                                StockTxn_3.ENTRY_LOC_ID = StockTxn_1.ENTRY_LOC_ID;
                                StockTxn_3.PERSONAL_DISC = (StockTxn_1.PERSONAL_DISC == "NULL") ? (item.PERSONAL_DISC == "NULL") ? "NULL" : item.PERSONAL_DISC : (item.PERSONAL_DISC == "NULL") ? StockTxn_1.PERSONAL_DISC : Convert.ToString(Convert.ToDouble(StockTxn_1.PERSONAL_DISC) + Convert.ToDouble(item.PERSONAL_DISC));
                                StockTxn_3.PENALTY = (StockTxn_1.PENALTY == "NULL") ? (item.PENALTY == "NULL") ? "NULL" : item.PENALTY : (item.PENALTY == "NULL") ? StockTxn_1.PENALTY : Convert.ToString(Convert.ToDouble(StockTxn_1.PENALTY) + Convert.ToDouble(item.PENALTY));
                                ExportTxnMeger(StockTxn_3);
                            }
                            else
                            {
                                ExportTxnError(StockTxn_1);
                                ExportTxnError(StockTxn_2);
                                ExportTxnError(item);
                            }

                            //下個迴圈前準備
                            StockTxn_1 = null;
                            StockTxn_2 = null;
                        }
                    }
                    else
                    {
                        //this.WriteLog(ExecutionMode.Debug, "E");
                        Console.WriteLine("E");
                        Thread.Sleep(100);
                        continue;
                    }
                }

                if (StockExportTxn_0 != null) { foreach (var item in StockExportTxn_0) { ExportTxnNormal(item); } }
                MyCommond.WriteLog(ThisReceive, $"判讀完成");
            }
            catch (Exception ex)
            {
                throw ex;
            }

        CalculatorEnd:
            CardTxnData = null;
            CardTxnData = new List<TxnTransaction>();
            Status = StatusEnum.Finish;

        }

        
        private void Calculator_CompareAnalyze()
        {

            Status = StatusEnum.Calculating;

            var sA = new string[] { "82", "205", "209" };

            CardList = (from p in MainServer.TxnAllData where sA.Contains(p.SVCE_LOC_ID) group p by new {p.CARD_PHYSICAL_ID} into g select g.Key.CARD_PHYSICAL_ID).ToList();

            Thread tt = new Thread(ExportTxn_CompareAnalyze);
            tt.IsBackground = true;

            foreach (var item in CardList)
            {
                this.Counts++;
                this.CardSN = item;
                if (!tt.IsAlive) { tt.Start(); }
                CardTxnData = (
                    from p in MainServer.TxnAllData 
                    where p.CARD_PHYSICAL_ID == item 
                    orderby p.TXN_TIMESTAMP
                    select p
                    ).ToList();
                if (CardTxnData.Count > 2) aaa();
            }

            bool ForWait = true;
            while (ForWait)
            {
                ForWait = (
                       ResponseExportTxnCompare.Count() > 0
                    || ResponseExportTxnMeger.Count() > 0
                    || ResponseExportTxnNormal.Count() > 0
                    || ResponseExportTxnError.Count() > 0
                    );
            }

        CalculatorEnd:
            CardTxnData.Clear();
            CardTxnData = null;
            Status = StatusEnum.Finish;
        } // 因2024/04/03大地震 影響板橋至中和段段軌，以接駁車載運，計算用

        private void aaa()
        {
            Console.WriteLine("New Card Txn");
            TxnTransaction StockTxn_1 = null;
            TxnTransaction StockTxn_2 = null;
            TxnTransaction StockTxn_3 = new TxnTransaction();
            try
            {
            ColculatorStart1:
                List<TxnTransaction> StockExportTxn_0 = CardTxnData.ToList();

                MyCommond.WriteLog(ThisReceive, $"判讀第{Counts, 8}張({CardSN, -20})");
                List<TxnTransaction> StockExportTxn_1 = new List<TxnTransaction>();  // 存放移除的資料
                List<TxnTransaction> StockExportTxn_2 = new List<TxnTransaction>();  // 存放合併的資料
                List<TxnTransaction> StockExportTxn_3 = new List<TxnTransaction>();  // 存放異常資料

                List<MrtStationList_Old> Stock_Mrt_Station = new List<MrtStationList_Old>();
                TimeSpan OverTime;

                OverTime = new TimeSpan(1, 0, 0);

                string[] TxnType_In = { "20" };
                string[] TxnType_Out = { "1", "4", "8", "13" };

                Stock_Mrt_Station.Add(new MrtStationList_Old() { CodeName = "82", StationName = "板橋", CodeName2 = "205", StationName2 = "中和" });
                Stock_Mrt_Station.Add(new MrtStationList_Old() { CodeName = "209", StationName = "板橋", CodeName2 = "205", StationName2 = "中和" });

                Stock_Mrt_Station.Add(new MrtStationList_Old() { CodeName = "205", StationName = "中和", CodeName2 = "209", StationName2 = "板橋" });
                Stock_Mrt_Station.Add(new MrtStationList_Old() { CodeName = "205", StationName = "中和", CodeName2 = "82", StationName2 = "板橋" });


                int StepCount = 0;

                foreach (var item in CardTxnData)
                {
                    if (StockExportTxn_0.Count() < 3) continue;
                    Console.WriteLine($"{StepCount++,4} - {item.ToString()}");
                    bool FindIn = TxnType_In.Contains(item.CARD_TXN_TYPE_ID);
                    bool FindOut = TxnType_Out.Contains(item.CARD_TXN_TYPE_ID);

                    bool FindStation_In = false;
                    bool FindStation_Out = false;

                    bool OnTime = false;

                    if (FindIn)
                    {
                        if (StockTxn_1 != null) OnTime = (Convert.ToDateTime(item.TXN_TIMESTAMP) - Convert.ToDateTime(StockTxn_1.TXN_TIMESTAMP)) <= OverTime;
                        FindStation_In = Stock_Mrt_Station.Find(x => x.CodeName == item.SVCE_LOC_ID) != null; 
                    }
                    if (FindOut)
                    {
                        FindStation_Out = Stock_Mrt_Station.Find(x => x.CodeName == item.SVCE_LOC_ID) != null;
                    }

                    Console.WriteLine($"進站 {FindIn}, 出站 {FindOut}, 進車站 {FindStation_In}, 出車站 {FindStation_Out}, 時間內 {OnTime}");

                    if (StockTxn_1 == null && FindOut && FindStation_Out)
                    {
                        StockTxn_1 = item;
                    }
                    else if (StockTxn_1 != null && StockTxn_2 == null && FindIn && FindStation_In && OnTime)
                    {
                        bool MaybeTransfare = Stock_Mrt_Station.Find(x => x.CodeName == item.SVCE_LOC_ID && x.CodeName2 == StockTxn_1.SVCE_LOC_ID) != null;
                        if (MaybeTransfare)
                        {
                            MyCommond.WriteLog(ThisReceive, $"疑似轉乘:");
                            MyCommond.WriteLog(ThisReceive, $"資料1:{StockTxn_1.ToString()}");
                            MyCommond.WriteLog(ThisReceive, $"資料2:{item.ToString()}");

                            ExportTxnCompare(StockTxn_1);
                            ExportTxnCompare(item);
                            StockTxn_2 = item;
                        }
                    }
                    else if (StockTxn_1 != null && StockTxn_2 != null)
                    {
                        MyCommond.WriteLog(ThisReceive, $"資料3:{item.ToString()}");
                        ExportTxnCompare(item);

                        StockTxn_3 = item;
                        StockTxn_3.TXN_AMT = Convert.ToString(Convert.ToInt16(StockTxn_1.TXN_AMT) + Convert.ToInt16(item.TXN_AMT));  //之後看到定期票要重算金額
                        StockTxn_3.ENTRY_LOC_ID = StockTxn_1.ENTRY_LOC_ID;
                        StockTxn_3.PERSONAL_DISC = 
                            (StockTxn_1.PERSONAL_DISC == "NULL") ? 
                                (item.PERSONAL_DISC == "NULL") ? 
                                    "NULL" 
                                    : item.PERSONAL_DISC 
                                : (item.PERSONAL_DISC == "NULL") ? 
                                    StockTxn_1.PERSONAL_DISC 
                                    : Convert.ToString(Convert.ToDouble(StockTxn_1.PERSONAL_DISC) + Convert.ToDouble(item.PERSONAL_DISC)
                            );
                        StockTxn_3.PENALTY = 
                            (StockTxn_1.PENALTY == "NULL") ? 
                                (item.PENALTY == "NULL") ? 
                                    "NULL" 
                                    : item.PENALTY 
                                : (item.PENALTY == "NULL") ? 
                                    StockTxn_1.PENALTY 
                                    : Convert.ToString(Convert.ToDouble(StockTxn_1.PENALTY) + Convert.ToDouble(item.PENALTY)
                            );

                        ExportTxnMeger(StockTxn_3);

                        StockTxn_1 = null;
                        StockTxn_2 = null;
                        StockTxn_3 = null;
                    }
                    else
                    {
                        Console.WriteLine("E");
                        continue;
                    }
                }

                StockExportTxn_1.Clear();
                StockExportTxn_2.Clear();
                StockExportTxn_3.Clear();
                Stock_Mrt_Station.Clear();
                StockTxn_1 = null;
                StockTxn_2 = null;
                StockTxn_3 = null;
                StockExportTxn_1 = null;
                StockExportTxn_2 = null;
                StockExportTxn_3 = null;
                Stock_Mrt_Station = null;

            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        private void GetCardTxn() => CardTxnData = MainServer.GetCardTxnData(ID.ToString(), CardSN);

        private void ExportTxn_CompareAnalyze()
        {
            bool Life_1 = true;
            bool Life_2 = true;
            bool Life_3 = true;
            bool Life_4 = true;
            bool Life_5 = true;
            bool Life_6 = true;

            while (Life_1 || Life_2 || Life_3 || Life_4 || Life_5 || Life_6)
            {
                Life_1 = MainServer.waitLoop_1;
                Life_2 = MainServer.waitLoop_2;
                Life_3 = ResponseExportTxnCompare.Count() > 0;
                Life_4 = ResponseExportTxnMeger.Count() > 0;
                Life_5 = ResponseExportTxnNormal.Count() > 0;
                Life_6 = ResponseExportTxnError.Count() > 0;

                if (Life_3)
                {
                    lock (ResponseExportTxnCompare)
                    {
                        var item = ResponseExportTxnCompare[0];
                        string path = $@"{MyCommond.Path_Program}Tidy\{MainServer.OperationDate:yyyy.MM}\TxnCompare{MainServer.OperationDate:yyyyMMdd}.csv";
                        try
                        {
                            if (!MyCommond.CheckFile(path)) MyCommond.ExportData(path, Encoding.UTF8, (new TxnTransaction()).ExportTitle_English());
                            MyCommond.ExportData(path, Encoding.UTF8, item.ToString());
                            ResponseExportTxnCompare.Remove(item);
                        }
                        catch (Exception ex)
                        {

                        }
                        finally
                        {

                        }
                    }
                }
                if (Life_4)
                {
                    lock (ResponseExportTxnMeger)
                    {
                        var item = ResponseExportTxnMeger[0];
                        string path = $@"{MyCommond.Path_Program}Tidy\{MainServer.OperationDate:yyyy.MM}\TxnMeger{MainServer.OperationDate:yyyyMMdd}.csv";
                        try
                        {
                            if (!MyCommond.CheckFile(path)) MyCommond.ExportData(path, Encoding.UTF8, (new TxnTransaction()).ExportTitle_English());
                            MyCommond.ExportData(path, Encoding.UTF8, item.ToString());
                            ResponseExportTxnMeger.Remove(item);
                        }
                        catch (Exception ex)
                        {

                        }
                        finally
                        {

                        }
                    }
                }
                if (Life_5)
                {
                    lock (ResponseExportTxnNormal)
                    {
                        var item = ResponseExportTxnNormal[0];
                        string path = $@"{MyCommond.Path_Program}Tidy\{MainServer.OperationDate:yyyy.MM}\TxnData{MainServer.OperationDate:yyyyMMdd}.csv";
                        try
                        {
                            if (!MyCommond.CheckFile(path)) MyCommond.ExportData(path, Encoding.UTF8, (new TxnTransaction()).ExportTitle_English());
                            MyCommond.ExportData(path, Encoding.UTF8, item.ToString());
                            ResponseExportTxnNormal.Remove(item);
                        }
                        catch (Exception ex)
                        {

                        }
                        finally
                        {

                        }
                    }
                }
                if (Life_6)
                {
                    lock (ResponseExportTxnError)
                    {
                        var item = ResponseExportTxnError[0];
                        string path = $@"{MyCommond.Path_Program}Tidy\{MainServer.OperationDate:yyyy.MM}\TxnError{MainServer.OperationDate:yyyyMMdd}.csv";
                        try
                        {
                            if (!MyCommond.CheckFile(path)) MyCommond.ExportData(path, Encoding.UTF8, (new TxnTransaction()).ExportTitle_English());
                            MyCommond.ExportData(path, Encoding.UTF8, item.ToString());
                            ResponseExportTxnError.Remove(item);
                        }
                        catch (Exception ex)
                        {

                        }
                        finally
                        {

                        }
                    }
                }
            }
        }

        private void ExportTxnNormal(TxnTransaction response)
        {
            lock (ResponseExportTxnNormal)
            {
                ResponseExportTxnNormal.Add(response);
            }
        }

        private void ExportTxnMeger(TxnTransaction response)
        {
            lock (ResponseExportTxnMeger)
            {
                ResponseExportTxnMeger.Add(response);
            }
        }

        private void ExportTxnError(TxnTransaction response)
        {
            lock (ResponseExportTxnError)
            {
                ResponseExportTxnError.Add(response);
            }
        }

        private void ExportTxnCompare(TxnTransaction response)
        {
            lock (ResponseExportTxnCompare)
            {
                ResponseExportTxnCompare.Add(response);
            }
        }

        #endregion

        #region 轉乘(月報表)

        private void Calculator_TrnAnalyze()
        {
            Status = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"轉程 - 開始");
            CardSN = MainServer.OperationDate.ToString("yyyyMMdd");
            int CountExit, ExitAmt, RideCountAdd = 0, RideCountMinus = 0, AmtAdd = 0, AmtMinus = 0;
            Int32 CountData = 0, CountRun = 0;
            DateTime OperationDate = MainServer.OperationDate;
            TimeSpan CTime;
            List<TxnTransaction> mTransaction = new List<TxnTransaction>();
            List<teatime> mTransactionMegre = new List<teatime>();
            DateTime StartTime = DateTime.Now;
            string FileName = $"history{OperationDate:yyyyMMdd}_nc.csv";

            string path_ = MyCommond.Path_Program + MyCommond.Path_Report + @"環狀線\轉乘統計\";

            string TidyfileName = MyCommond.CheckAndReName(path_, $"{OperationDate:yyyyMMdd}_Trn_GroupData", "csv");
            string[] Data = null;

            MyCommond.WriteLog(ThisReceive, $"分析檔案為 {FileName}");
            MyCommond.WriteLog(ThisReceive, $"資料分析後儲存為 {path_}{TidyfileName}");
            try
            {
                var List_Compare_TxnType = MainServer.globalCommond.Stock_Mrt_TxnType.FindAll(x => x.Num == 1 || x.Num == 4 || x.Num == 13).ToList();
                string[] Array_Compare_TxnType = new string[List_Compare_TxnType.Count];
                for (int i = 0; i < List_Compare_TxnType.Count; i++) { Array_Compare_TxnType[i] = List_Compare_TxnType[i].ChName; }

                if (!true) 
                { 
                    #region 資料載入

                if (MyCommond.CheckFile(MyCommond.Path_Program + MyCommond.Path_Analyze + FileName))
                {
                    Data = File.ReadAllLines(MyCommond.Path_Program + MyCommond.Path_Analyze + FileName);
                    CountData = Data.Length;

                    foreach (var item in Data)
                    {
                        CountRun++;
                        if (CountRun == 1) continue; //跳過表頭
                        TxnTransaction ListTemp = new TxnTransaction();
                        var jtem = item.Split(',');
                        string msg1 = $"{CountRun,8}/{CountData} CARD_TXN_TYPE_ID:{jtem[0],4}, SVCE_LOC_ID:{jtem[9],4}, BUSINESS_DATE:{Convert.ToDateTime(jtem[11]):yyyy-MM-dd}";
                        string msg2 = "";
                        int ColumnCount = 0;
                        Type type = typeof(TxnTransaction);
                        if (Convert.ToInt16(jtem[9]) >= 200)
                        {
                            foreach (var mTemp in jtem)
                            {
                                type.GetProperties()[ColumnCount].SetValue(ListTemp, Convert.ToString(mTemp), null);
                                ColumnCount++;
                            }
                            mTransaction.Add(ListTemp);
                        }
                        MyCommond.WriteLog(ThisReceive, $"{msg1 + msg2}");
                    }
                }

                #endregion
                
                    mTransactionMegre = (
                        from p in mTransaction
                        where p.CARD_TXN_TYPE_ID == "1" || p.CARD_TXN_TYPE_ID == "4" || p.CARD_TXN_TYPE_ID == "13"
                        group p by new { p.CARD_TXN_SUBTYPE_ID, p.FARE_PRODUCT_TYPE_ID, p.ISSUER_ID, p.XFER_CODE, p.BUSINESS_DATE, p.CARD_TXN_TYPE_ID, p.SVCE_LOC_ID, p.ENTRY_LOC_ID } into g
                        select new teatime
                        {
                            BUSINESS_DATE = g.Key.BUSINESS_DATE,
                            CARD_TXN_SUBTYPE_ID = g.Key.CARD_TXN_SUBTYPE_ID,
                            CARD_TXN_TYPE_ID = g.Key.CARD_TXN_TYPE_ID,
                            ENTRY_LOC_ID = g.Key.ENTRY_LOC_ID,
                            SVCE_LOC_ID = g.Key.SVCE_LOC_ID,
                            ISSUER_ID = g.Key.ISSUER_ID,
                            XFER_CODE = g.Key.XFER_CODE,
                            FARE_PRODUCT_TYPE_ID = g.Key.FARE_PRODUCT_TYPE_ID,
                            Counts = g.Count(),
                            Amts = g.Sum(p => (int)Convert.ToDouble((p.XFER_DISC != "NULL") ? p.XFER_DISC : "0"))
                        }).ToList();
                }
                else
                {
                    mTransactionMegre = (
                        from p in MainServer.Data_Merge0
                        where Array_Compare_TxnType.Contains(p.CARD_TXN_TYPE_ID)
                        group p by new { p.CARD_TXN_SUBTYPE_ID, p.FARE_PRODUCT_TYPE_ID, p.ISSUER_ID, p.XFER_CODE, p.BUSINESS_DATE, p.CARD_TXN_TYPE_ID, p.SVCE_LOC_ID, p.ENTRY_LOC_ID } into g
                        select new teatime
                        {
                            BUSINESS_DATE = g.Key.BUSINESS_DATE,
                            CARD_TXN_SUBTYPE_ID = g.Key.CARD_TXN_SUBTYPE_ID,
                            CARD_TXN_TYPE_ID = g.Key.CARD_TXN_TYPE_ID,
                            ENTRY_LOC_ID = g.Key.ENTRY_LOC_ID,
                            SVCE_LOC_ID = g.Key.SVCE_LOC_ID,
                            ISSUER_ID = g.Key.ISSUER_ID,
                            XFER_CODE = g.Key.XFER_CODE,
                            FARE_PRODUCT_TYPE_ID = g.Key.FARE_PRODUCT_TYPE_ID,
                            Counts = g.Count(),
                            Amts = g.Sum(p => (int)Convert.ToDouble((p.XFER_DISC != "NULL") ? p.XFER_DISC : "0"))
                        }).ToList();
                }

                MyCommond.ExportData(path_ + TidyfileName, Encoding.UTF8, (new teatime()).ExportTitle_zhTW());
                foreach (var item in mTransactionMegre) { MyCommond.ExportData(path_ + TidyfileName, Encoding.UTF8, item.ToString()); }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        ThisEnd:
            DateTime EndTime = DateTime.Now;
            CTime = EndTime - StartTime;
            CountExit = RideCountAdd - RideCountMinus;
            ExitAmt = AmtAdd - AmtMinus;
            MyCommond.Path_Analyze = @"Analyze\";
            GC.Collect();
            MyCommond.WriteLog(ThisReceive, $"轉程 - 結束");
            Status = StatusEnum.Finish;
        }

        #endregion

        #region 通勤月票(月報表)

        private void Calculator_TPassAnalyze()
        {
            Status = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"通勤月報 - 開始");
            CardSN = MainServer.OperationDate.ToString("yyyyMMdd");
            int CountExit, ExitAmt, RideCountAdd = 0, RideCountMinus = 0, AmtAdd = 0, AmtMinus = 0;
            Int32 CountData = 0, CountRun = 0;
            DateTime OperationDate = MainServer.OperationDate;
            TimeSpan CTime;
            string path_ = MyCommond.Path_Program + MyCommond.Path_Report + @"環狀線\通勤月票明細\";
            DateTime StartTime = DateTime.Now;
            string FileName = $"history{OperationDate:yyyyMMdd}_nc.csv";
            string[] Data = null;

            CountData = 0;
            CountRun = 0;
            RideCountAdd = 0;
            RideCountMinus = 0;
            AmtAdd = 0;
            AmtMinus = 0;

            try
            {
                #region 資料載入

                List<TxnTransaction> mTransaction = new List<TxnTransaction>();

                if (MyCommond.CheckFile(MyCommond.Path_Program + MyCommond.Path_Analyze + FileName))
                {
                    MyCommond.WriteLog(ThisReceive, $"取得交易資料");
                    Data = File.ReadAllLines(MyCommond.Path_Program + MyCommond.Path_Analyze + FileName);
                    CountData = Data.Length;

                    foreach (var item in Data)
                    {
                        CountRun++;
                        if (CountRun == 1) continue; //跳過表頭
                        TxnTransaction ListTemp = new TxnTransaction();
                        var jtem = item.Split(',');
                        string msg1 = $"{CountRun,8}/{CountData} CARD_TXN_TYPE_ID:{jtem[0],4}, SVCE_LOC_ID:{jtem[9],4}, BUSINESS_DATE:{Convert.ToDateTime(jtem[11]):yyyy-MM-dd}";
                        string msg2 = "";
                        if (Convert.ToDateTime(jtem[11]) >= Convert.ToDateTime(OperationDate.ToString("yyyy/MM/dd 00:00:00")) && Convert.ToDateTime(jtem[11]) < Convert.ToDateTime(OperationDate.AddDays(1).ToString("yyyy/MM/dd 00:00:00")))
                        {
                            if (Convert.ToInt16(jtem[9]) >= 200)
                            {
                                int ColumnCount = 0;
                                Type type = typeof(TxnTransaction);

                                foreach (var mTemp in jtem)
                                {
                                    type.GetProperties()[ColumnCount].SetValue(ListTemp, Convert.ToString(mTemp), null);
                                    ColumnCount++;
                                }
                                mTransaction.Add(ListTemp);
                            }
                            else { msg2 = $", Station Comparison False"; }
                        }
                        else { msg2 = $",    Time Comparison False"; }
                        MyCommond.WriteLog(ThisReceive, msg1 + msg2);
                    }
                }

                #endregion

                MyCommond.WriteLog(ThisReceive, $"整理資料");

                if (!(MyCommond.CheckFile(path_ + OperationDate.ToString("yyyy-MM") + "_1200設定.csv"))) MyCommond.ExportData(path_ + OperationDate.ToString("yyyy-MM") + "_1200設定.csv", Encoding.UTF8, (new TxnTransaction()).ExportTitle_English());
                if (!(MyCommond.CheckFile(path_ + OperationDate.ToString("yyyy-MM") + "_1200退票.csv"))) MyCommond.ExportData(path_ + OperationDate.ToString("yyyy-MM") + "_1200退票.csv", Encoding.UTF8, (new TxnTransaction()).ExportTitle_English());
                if (!(MyCommond.CheckFile(path_ + OperationDate.ToString("yyyy-MM") + "_1200取消.csv"))) MyCommond.ExportData(path_ + OperationDate.ToString("yyyy-MM") + "_1200取消.csv", Encoding.UTF8, (new TxnTransaction()).ExportTitle_English());

                List<TxnTransaction> TxnType_Set =
                    (from p in mTransaction
                     where p.CARD_TXN_TYPE_ID == "3" && p.FARE_PRODUCT_TYPE_ID == "75"
                     select p).ToList();
                foreach (var item in TxnType_Set)
                {
                    MyCommond.ExportData(path_ + OperationDate.ToString("yyyy-MM") + "_1200設定.csv", Encoding.UTF8, item.ToString());
                }

                List<TxnTransaction> TxnType_Refund =
                    (from p in mTransaction
                     where p.CARD_TXN_TYPE_ID == "6" && p.FARE_PRODUCT_TYPE_ID == "75"
                     select p).ToList();
                foreach (var item in TxnType_Refund)
                {
                    MyCommond.ExportData(path_ + OperationDate.ToString("yyyy-MM") + "_1200退票.csv", Encoding.UTF8, item.ToString());
                }

                List<TxnTransaction> TxnType_Cancel =
                    (from p in mTransaction
                     where p.CARD_TXN_TYPE_ID == "10" && p.FARE_PRODUCT_TYPE_ID == "75"
                     select p).ToList();
                foreach (var item in TxnType_Cancel)
                {
                    MyCommond.ExportData(path_ + OperationDate.ToString("yyyy-MM") + "_1200取消.csv", Encoding.UTF8, item.ToString());
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            DateTime EndTime = DateTime.Now;
            CTime = EndTime - StartTime;
            CountExit = RideCountAdd - RideCountMinus;
            ExitAmt = AmtAdd - AmtMinus;
            GC.Collect();
            MyCommond.WriteLog(ThisReceive, $"通勤月報 - 結束");
            Status = StatusEnum.Finish;
        }

        #endregion

        #region 運量報表

        private string ExportPath;
        private string ExportFileName;
        private List<LrtTxnClientSwitch> ReportMethode = new List<LrtTxnClientSwitch>();
        private void Calculator_VolumeReportAnalyze()
        {
            Status = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"運量報表產製 - 開始");
            CardSN = MainServer.OperationDate.ToString("yyyyMMdd");
            goto StartReport;
            

        StartReport:
            List<Thread> threads = new List<Thread>();

            foreach (var item in MainServer.List_ReportBox)
            {
                Thread thread = new Thread(() =>
                {
                    try { 
                             if (item.methode == (LrtTxnClientSwitch)1 ) { Calculator_Volume(); }
                        else if (item.methode == (LrtTxnClientSwitch)2 ) { Calculator_Amount(); }
                        else if (item.methode == (LrtTxnClientSwitch)3 ) { Calculator_EachStationEachTime(); }
                        else if (item.methode == (LrtTxnClientSwitch)4 ) { Calculator_ElectronicTicket(); }
                        else if (item.methode == (LrtTxnClientSwitch)5 ) { /* 日運量統計表 */ }
                        else if (item.methode == (LrtTxnClientSwitch)6 ) { Calculator_TrafficAmount(); }
                        else if (item.methode == (LrtTxnClientSwitch)7 ) { Calculator_OriginDestination(); }
                        else if (item.methode == (LrtTxnClientSwitch)8 ) { Calculator_EquipAmount(); }
                        else if (item.methode == (LrtTxnClientSwitch)9 ) { Calculator_Month_AllRideList(); }
                        else if (item.methode == (LrtTxnClientSwitch)10) { Calculator_Month_OriginDestination(); }
                        else if (item.methode == (LrtTxnClientSwitch)11) { Calculator_Month_ElectronicTicket_Station(); }
                        else if (item.methode == (LrtTxnClientSwitch)12) { Calculator_Month_ElectronicTicket_Day(); }
                        else if (item.methode == (LrtTxnClientSwitch)13) { Calculator_Month_OwnTicketVolume_Station(); }
                        else if (item.methode == (LrtTxnClientSwitch)14) { Calculator_Month_OwnTicketVolume_Day(); }
                        else if (item.methode == (LrtTxnClientSwitch)15) { Calculator_Month_TrafficAmount_Station(); }
                        else if (item.methode == (LrtTxnClientSwitch)16) { Calculator_Month_TrafficAmount_Day(); }
                        else if (item.methode == (LrtTxnClientSwitch)17) { Calculator_Month_EquipAmount_Station(); }
                        else if (item.methode == (LrtTxnClientSwitch)18) { Calculator_Month_EquipAmount_Day(); }
                        else if (item.methode == (LrtTxnClientSwitch)19) { Calculator_Month_Volumn(); }
                        else if (item.methode == (LrtTxnClientSwitch)20) { Calculator_Month_Amount(); }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                });
                thread.Start();
                threads.Add(thread);
            }

            foreach (var item in threads)
            {
                item.Join();
            }

            MyCommond.WriteLog(ThisReceive, $"運量報表產製 - 結束");
            Status = StatusEnum.Finish;
        }

        #region 日

        private void Calculator_Volume()
        {
            var MethodeName = LrtTxnClientSwitch.Day_Volume;
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");

            var reportTemp = MainServer.globalCommond.Stock_ReportFile.Find(x => 
                x.ReportNameEng == MethodeName.ToString()
                && x.OperationLine == MainServer.globalCommond.OperationCode
                );
            int cs = MyCommond.ExcelColumnNum(reportTemp.StartColumn);
            int ce = MyCommond.ExcelColumnNum(reportTemp.EndColumn);

            double[,] rd = new double[1, 1 + ce - cs];
            string[] fdsf = new string[] { "0x1012", "0x1022", "0x1111", "0x1121" /*, "0x1002", "0x1102", "0x1112", "0x1122"*/ };
            var ColumnString = MainServer.globalCommond.Stock_Mrt_TicketType.FindAll(x => x.Check == "o").ToList();

            try
            {
                // 轉換 Remark 16進位變10進位 後大於 255 且 符合上述4個篩選
                var data_merge0_1 = (
                    from p in MainServer.Data_Merge0 
                    where 
                        Int32.Parse(p.Remark.Substring(2, p.Remark.Length - 2), System.Globalization.NumberStyles.HexNumber) > 255 
                        && fdsf.Contains(p.Remark) 
                    group p by new {p.TicketType} into g
                    select new { g.Key.TicketType, Counts = g.Count()}
                    ).ToList();

                foreach (var item in data_merge0_1)
                {
                    try
                    {
                        int offsetColumn = -1;
                        offsetColumn = Convert.ToInt16(ColumnString.Find(x => x.TicketName == item.TicketType).Volume) - 1;
                        if (offsetColumn < 0) continue;
                        rd[0, offsetColumn] += item.Counts;
                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $"資料異常(非本次要計算的資料):{item.ToString()}");
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            string exportpath = MyCommond.Path_Program 
                + MyCommond.Path_Report 
                + @"環狀線\" 
                + string.Format(MyCommond.Path_Date, MainServer.OperationDate.ToString("yyyy.MM.dd"));
            string exportfilename = MyCommond.CheckAndReName(exportpath, new ExportConfig()
            {
                Path1 = exportpath,
                DTime = MainServer.OperationDate,
                yearType = YearType.RC,
                FirstName = "",
                FileName = $"{MainServer.globalCommond.Stock_ReportFile.Find(x => x.ReportNameEng == MethodeName.ToString() && x.OperationLine == MainServer.globalCommond.OperationCode).ReportNameCh}",
                SubName = "",
                OpLine = MainServer.globalCommond.OperationCode_String,
                LastName = "xlsx",
            }.ToString(), "xlsx");
            ExportExcle(new ExportConfig
            {
                Path1 = exportpath,
                FileName = exportfilename,
            }, MethodeName, rd, null, null, null);

            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 結束");
        } // 日 - 運量(完成)

        private void Calculator_Amount()
        {
            var MethodeName = LrtTxnClientSwitch.Day_Amount;
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");

            double _amt = 0;
            System.Threading.Thread thread = new Thread(() =>
            {
                var a1 = (from p in MainServer.Data_Merge0 group p by new { p.M_Entry, p.M_Exit } into g select new { g.Key.M_Entry, g.Key.M_Exit, mSum = g.Sum(x => Convert.ToDouble(x.TicketSubsidy)) }).ToList();

                _amt = a1.Sum(x => Math.Round(x.mSum));
            });
            thread.IsBackground = true;
            thread.Start();

            var reportTemp = MainServer.globalCommond.Stock_ReportFile.Find(x =>
                x.ReportNameEng == MethodeName.ToString()
                && x.OperationLine == MainServer.globalCommond.OperationCode
                );
            int cs = MyCommond.ExcelColumnNum(reportTemp.StartColumn);
            int ce = MyCommond.ExcelColumnNum(reportTemp.EndColumn);

            double[,] rd = new double[1, 1 + ce - cs];
            string[] fdsf = new string[] { "0x1012", "0x1022", "0x1111", "0x1121" /*, "0x1002", "0x1102", "0x1112", "0x1122"*/ };
            var ColumnString = MainServer.globalCommond.Stock_Mrt_TicketType.FindAll(x => x.Check == "o").ToList();


            try
            {
                // 轉換 Remark 16進位變10進位 後大於 255 且 符合上述4個篩選
                var data_merge0_1 = (
                    from p in MainServer.Data_Merge0
                    where
                        Int32.Parse(p.Remark.Substring(2, p.Remark.Length - 2), System.Globalization.NumberStyles.HexNumber) > 255
                        && fdsf.Contains(p.Remark)
                    select p
                    ).ToList();

                foreach (var item in data_merge0_1)
                {
                    try
                    {
                        int offsetColumn = -1;
                        offsetColumn = Convert.ToInt16(ColumnString.Find(x => x.TicketName == item.TicketType).Amount) - 1;
                        if (offsetColumn < 0) continue;
                        rd[0, offsetColumn] += Convert.ToDouble(item.RealPay);
                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $"資料異常(非本次要計算的資料):{item.ToString()}");
                    }
                }

                int cco = Convert.ToInt16(ColumnString.Max(x => x.Amount));
                /*--社福及轉乘--*/ rd[0, cco++] = MainServer.Data_Merge0.Sum(x => Convert.ToDouble(x.Split_DiscAmount))
                                            + MainServer.Data_Merge0.Sum(x => Convert.ToDouble(x.Split_TrnNormalAmount));
                /*-公共運輸票收-*/ rd[0, cco++] = data_merge0_1.Sum(x => Convert.ToDouble(x.Split_AllPassAmount));
                thread.Join();
                /*---運價票差---*/ rd[0, cco++] = _amt;

            }
            catch (Exception ex)
            {
                throw ex;
            }


            goto jjmp;

            /*---電子票證---*/ var C_1 = (from p in MainServer.globalCommond.Stock_Mrt_TicketType where p.Check == "o" && p.Amount == "1" select p.TicketName).ToList();
            /*---行動支付---*/ var C_2 = (from p in MainServer.globalCommond.Stock_Mrt_TicketType where p.Check == "o" && p.Amount == "2" select p.TicketName).ToList();
            /*----單程票----*/ var C_3 = (from p in MainServer.globalCommond.Stock_Mrt_TicketType where p.Check == "o" && p.Amount == "3" select p.TicketName).ToList();
            /*----天期票----*/ var C_4 = (from p in MainServer.globalCommond.Stock_Mrt_TicketType where p.Check == "o" && p.Amount == "4" select p.TicketName).ToList();
            /*----小時票----*/ var C_5 = (from p in MainServer.globalCommond.Stock_Mrt_TicketType where p.Check == "o" && p.Amount == "5" select p.TicketName).ToList();

            try
            {
                var Csc_Card = (
                    from p in MainServer.Data_Merge0
                    where
                        Int32.Parse(p.Remark.Substring(2, p.Remark.Length - 2), System.Globalization.NumberStyles.HexNumber) > 255
                        && fdsf.Contains(p.Remark)
                    select p
                    ).ToList();
                int c_Offset = 0;
                /*---電子票證---*/ rd[0, c_Offset++] = (from p in Csc_Card where C_1.Contains(p.TicketType) select p).ToList().Sum(x => Convert.ToDouble(x.RealPay));
                /*---行動支付---*/ rd[0, c_Offset++] = (from p in Csc_Card where C_2.Contains(p.TicketType) select p).ToList().Sum(x => Convert.ToDouble(x.RealPay));
                /*----單程票----*/ rd[0, c_Offset++] = (from p in Csc_Card where C_3.Contains(p.TicketType) select p).ToList().Sum(x => Convert.ToDouble(x.RealPay));
                /*----天期票----*/ rd[0, c_Offset++] = (from p in Csc_Card where C_4.Contains(p.TicketType) select p).ToList().Sum(x => Convert.ToDouble(x.RealPay));
                /*----小時票----*/ rd[0, c_Offset++] = (from p in Csc_Card where C_5.Contains(p.TicketType) select p).ToList().Sum(x => Convert.ToDouble(x.RealPay));
                /*--社福及轉乘--*/ rd[0, c_Offset++] = MainServer.Data_Merge0.Sum(x => Convert.ToDouble(x.Split_DiscAmount))
                                            + MainServer.Data_Merge0.Sum(x => Convert.ToDouble(x.Split_TrnNormalAmount));
                /*-公共運輸票收-*/ rd[0, c_Offset++] = Csc_Card.Sum(x => Convert.ToDouble(x.Split_AllPassAmount));
                thread.Join();
                /*---運價票差---*/ rd[0, c_Offset++] = _amt;

            }
            catch (Exception ex)
            {
                throw ex;
            }

            jjmp:
            string exportpath = MyCommond.Path_Program
                + MyCommond.Path_Report
                + @"環狀線\"
                + string.Format(MyCommond.Path_Date, MainServer.OperationDate.ToString("yyyy.MM.dd"));
            string exportfilename = MyCommond.CheckAndReName(exportpath, new ExportConfig()
            {
                Path1 = exportpath,
                DTime = MainServer.OperationDate,
                yearType = YearType.RC,
                FirstName = "",
                FileName = $"{MainServer.globalCommond.Stock_ReportFile.Find(x => x.ReportNameEng == MethodeName.ToString() && x.OperationLine == MainServer.globalCommond.OperationCode).ReportNameCh}",
                SubName = "",
                OpLine = MainServer.globalCommond.OperationCode_String,
                LastName = "xlsx",
            }.ToString(), "xlsx");
            ExportExcle(new ExportConfig
            {
                Path1 = exportpath,
                FileName = exportfilename,
            }, MethodeName, rd, null, null, null);

            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 結束");
        } // 日 - 營收(完成)

        private void Calculator_EachStationEachTime()
        {
            var MethodeName = LrtTxnClientSwitch.Day_EachStation_EachTime;
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");
            var stationCounts = MainServer.globalCommond.Stock_Mrt_Y_Station.FindAll(x => x.Company == "新北捷運").ToList();

            var mYear = ((Convert.ToInt32(MainServer.OperationDate.ToString("yyyy")) > 1911) ? (Convert.ToInt32(MainServer.OperationDate.ToString("yyyy")) - 1911) : Convert.ToInt32(MainServer.OperationDate.ToString("yyyy"))).ToString();
            var mMonth = MainServer.OperationDate.ToString("MM");
            var mDay = MainServer.OperationDate.ToString("dd");

            string fileName1 = $"{mYear}{mMonth}{mDay}";

            string tempfilepath = MyCommond.Path_Program + MyCommond.Path_Template;
            string tempfile = MethodeName.ToString();
            string exportpath = MyCommond.Path_Program + MyCommond.Path_Report + @"環狀線\" + string.Format(MyCommond.Path_Date, MainServer.OperationDate.ToString("yyyy.MM.dd"));
            string exportfilename = MyCommond.CheckAndReName(exportpath, new ExportConfig()
            {
                Path1 = exportpath,
                DTime = MainServer.OperationDate,
                yearType = YearType.RC,
                FirstName = "",
                FileName = $"{MainServer.globalCommond.Stock_ReportFile.Find(x => x.ReportNameEng == MethodeName.ToString() && x.OperationLine == MainServer.globalCommond.OperationCode).ReportNameCh}",
                SubName = "",
                OpLine = MainServer.globalCommond.OperationCode_String,
                LastName = "xlsx",
            }.ToString(), "xlsx");

            int temp_Day_Hour = 24;
            int temp_Start_Hour = 5;
            string[,] data = new string[24, stationCounts.Count * 2];
            int RideCounts = 0;

            MyCommond.WriteLog(ThisReceive, $"aaaaa");
            string[] fdsf = new string[] { "0x1012", "0x1022", "0x1111", "0x1121" /*, "0x1002", "0x1102", "0x1112", "0x1122"*/ };
            var data_merge0_1 = (from p in MainServer.Data_Merge0 where Int32.Parse(p.Remark.Substring(2, p.Remark.Length - 2), System.Globalization.NumberStyles.HexNumber) > 255 && (fdsf.Contains(p.Remark) /*&& asdf.Contains(p.SVCE_LOC_ID)*/) select p).ToList();

            List<TxnTransactionOd> MetroTrnNtmetro = new List<TxnTransactionOd>();
            double[] trn = new double[stationCounts.Count * 2];
            double[,] data2 = null;
            try 
            { 
                foreach (var item in data_merge0_1)
                {
                    #region 各站分時主體
                    string[] trnStationAll = new string[] { "大坪林", "Y大坪林", "景安", "Y景安", "板橋", "Y板橋", "頭前庄", "Y頭前庄", "新埔", "新埔民生" };
                         if (trnStationAll.Contains(item.M_Entry)) { int temp = 0; }
                    else if (trnStationAll.Contains(item.M_Exit))  { int temp = 0; }

                    var temp_Hour_Entry = Convert.ToInt16(Convert.ToDateTime(item.ENTRY_DATETIME).ToString("HH"));
                    int Offset_Row_Entry = (temp_Hour_Entry - temp_Start_Hour >= 0) ? temp_Hour_Entry - temp_Start_Hour : temp_Day_Hour + temp_Hour_Entry - temp_Start_Hour;
                    int Offset_Column_Entry = data.GetLength(0) - 1;

                    if (item.M_Entry != "NULL") { Offset_Column_Entry = stationCounts.FindIndex(x => x.StationName == item.M_Entry); }
                    if (Offset_Column_Entry == -1) { Offset_Column_Entry = stationCounts.FindIndex(x => x.StationName == $"Y{item.M_Entry}"); }
                    if (Offset_Column_Entry == -1) { Offset_Column_Entry = stationCounts.FindIndex(x => x.StationName == $"{item.M_Entry}民生"); }
                    if (Offset_Column_Entry != -1)
                    {

                    }

                    RideCounts++;

                    var temp_Hour_Exit_ = Convert.ToInt16(Convert.ToDateTime(item.TXN_TIMESTAMP).ToString("HH"));
                    var Offset_Row_Exit_ = (temp_Hour_Exit_ - temp_Start_Hour >= 0) ? temp_Hour_Exit_ - temp_Start_Hour : temp_Day_Hour + temp_Hour_Exit_ - temp_Start_Hour;
                    int Offset_Column_Exit_ = data.GetLength(1) - 1;

                    if (item.M_Exit != "NULL") Offset_Column_Exit_ = stationCounts.FindIndex(x => x.StationName == item.M_Exit);
                    if (Offset_Column_Exit_ == -1) Offset_Column_Exit_ = stationCounts.FindIndex(x => x.StationName == $"Y{item.M_Exit}");
                    if (Offset_Column_Exit_ == -1) Offset_Column_Exit_ = stationCounts.FindIndex(x => x.StationName == $"{item.M_Exit}民生");
                    if (Offset_Column_Entry != -1)
                    {

                    }

                    if (Offset_Column_Entry == -1 || Offset_Column_Exit_ == -1)
                    {
                        MyCommond.WriteLog(ThisReceive, $"Error! EntryIndex:{Offset_Column_Entry,4}, ExitIndex:{Offset_Column_Exit_,4}");
                        continue;
                    }

                    data[Offset_Row_Entry, Offset_Column_Entry * 2 + 0] = (Convert.ToInt32(data[Offset_Row_Entry, Offset_Column_Entry * 2 + 0]) + 1).ToString();
                    data[Offset_Row_Exit_, Offset_Column_Exit_ * 2 + 1] = (Convert.ToInt32(data[Offset_Row_Exit_, Offset_Column_Exit_ * 2 + 1]) + 1).ToString();
                    #endregion
                    #region 環狀線轉程量計算
                    bool fsd = false;
                    string[] TrnStationSmall_1 = new string[] { "大坪林", "Y大坪林", "景安", "Y景安", "頭前庄", "Y頭前庄" };
                    if (TrnStationSmall_1.Contains(item.M_Entry))
                    {
                        if (!fsd) { fsd = $"{item.M_Entry}" == $"{item.ENTRY_LOC_ID}"; }
                        if (!fsd) { fsd = $"{item.M_Entry}" == $"Y{item.ENTRY_LOC_ID}"; }
                        if (!fsd) { fsd = $"Y{item.M_Entry}" == $"{item.ENTRY_LOC_ID}"; }
                        if (!fsd) { fsd = $"{item.M_Entry}" == $"{item.ENTRY_LOC_ID}民生"; }
                        if (!fsd) { fsd = $"民生{item.M_Entry}" == $"{item.ENTRY_LOC_ID}"; }

                        if (!fsd) 
                        {
                            MetroTrnNtmetro.Add(item);
                            trn[Offset_Column_Entry * 2 + 0]++;
                        }
                    }
                    #endregion
                }

                //輸出環傳線轉乘輛明細
                if (!true) MyCommond.ExportData<TxnTransactionOd>(ThisReceive, new ExportConfig()
                {
                    Path1 = MyCommond.Path_Program + MyCommond.Path_Report,
                    DTime = MainServer.OperationDate,
                    yearType = YearType.RC,
                    FirstName = "",
                    FileName = "非三站進",
                    SubName = "",
                    OpLine = "",
                    LastName = "xlsx",
                }, new TxnTransactionOd().ExportTitle_zhTW(), MetroTrnNtmetro, Encoding.UTF8);

                data2 = new double[data.GetLength(0), data.GetLength(1)];
                for (int i = 0; i < data.GetLength(0); i++)
                {
                    for (int j = 0; j < data.GetLength(1); j++)
                    {
                        data2[i, j] = Convert.ToInt32(data[i, j]);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            //輸出報表
            ExportExcle(new ExportConfig()
            {
                Path1 = exportpath,
                DTime = new DateTime(),
                yearType = YearType.RC,
                FirstName = "",
                FileName = exportfilename,
                SubName = "",
                OpLine = "",
                LastName = "",
            }, MethodeName, data2, null, null, trn);
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 結束");
        } // 日 - 分時各站(完成)

        private void Calculator_ElectronicTicket()
        {
            var MethodeName = LrtTxnClientSwitch.Day_ElectronicTicket;
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");
            var stationCounts = MainServer.globalCommond.Stock_Mrt_Y_Station.FindAll(x => x.Company == "新北捷運").ToList();
            var ColumnString = MainServer.globalCommond.Stock_Mrt_TicketType.FindAll(x => x.Check == "o").ToList();
            double[,] rd = new double[stationCounts.Count, ColumnString.Count * 2];

            string[] fdsf = new string[] { "0x1012", "0x1022", "0x1111", "0x1121" /*, "0x1002", "0x1102", "0x1112", "0x1122"*/ };
            // 轉換 Remark 16進位變10進位 後大於 255 且 符合上述4個篩選
            var data_merge0_1 = (from p in MainServer.Data_Merge0 where Int32.Parse(p.Remark.Substring(2, p.Remark.Length - 2), System.Globalization.NumberStyles.HexNumber) > 255 && (fdsf.Contains(p.Remark)) select p).ToList();
            try
            {
                foreach (var item in data_merge0_1)
                {
                    int Offset_Column = -1;
                    try
                    {
                        Offset_Column = ColumnString.FindIndex(x => x.TicketName == item.TicketType);

                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $"ERROR\nOffset_Column:{Offset_Column}\nitem:{item.ToString()}");
                    }

                    int Offset_Row_Entry = -1;
                    try
                    {
                        if (item.M_Entry != "NULL") { Offset_Row_Entry = stationCounts.FindIndex(x => x.StationName == item.M_Entry); }
                        if (Offset_Row_Entry == -1) { Offset_Row_Entry = stationCounts.FindIndex(x => x.StationName == $"Y{item.M_Entry}"); }
                        if (Offset_Row_Entry == -1) { Offset_Row_Entry = stationCounts.FindIndex(x => x.StationName == $"{item.M_Entry}民生"); }
                        if (Offset_Row_Entry != -1) { }

                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $"ERROR\nOffset_Row_Entry:{Offset_Row_Entry}\nitem:{item.ToString()}");
                    }

                    int Offset_Row_Exit = -1;
                    try
                    {
                        if (item.M_Entry != "NULL") { Offset_Row_Exit = stationCounts.FindIndex(x => x.StationName == item.M_Exit); }
                        if (Offset_Row_Exit == -1) { Offset_Row_Exit = stationCounts.FindIndex(x => x.StationName == $"Y{item.M_Exit}"); }
                        if (Offset_Row_Exit == -1) { Offset_Row_Exit = stationCounts.FindIndex(x => x.StationName == $"{item.M_Exit}民生"); }
                        if (Offset_Row_Exit != -1) { }
                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $"ERROR\nOffset_Row_Exit:{Offset_Row_Exit}\nitem:{item.ToString()}");
                    }

                    if (Offset_Column > -1 && Offset_Row_Entry > -1 && Offset_Row_Exit > -1)
                    {
                        rd[Offset_Row_Entry, Offset_Column * 2 + 0]++;
                        rd[Offset_Row_Exit , Offset_Column * 2 + 1]++;
                    }
                }
                /// if (!true) { MyCommond.ExportData(ThisReceive, new ExportConfig()
                ///             {
                ///                 Path1 = MyCommond.Path_Program + MyCommond.Path_Report,
                ///                 DTime = MainServer.OperationDate,
                ///                 FileName = "票卡進出測試",
                ///                 OpLine = MainServer.globalCommond.OperationLine_String,
                ///                 LastName = "csv",
                ///             }, null, rd, Encoding.UTF8); }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            string exportpath = MyCommond.Path_Program + MyCommond.Path_Report + @"環狀線\" + string.Format(MyCommond.Path_Date, MainServer.OperationDate.ToString("yyyy.MM.dd"));
            string exportfilename = MyCommond.CheckAndReName(exportpath, new ExportConfig()
            {
                Path1 = exportpath,
                DTime = MainServer.OperationDate,
                yearType = YearType.RC,
                FirstName = "",
                FileName = $"{MainServer.globalCommond.Stock_ReportFile.Find(x => x.ReportNameEng == MethodeName.ToString()).ReportNameCh}",
                SubName = "",
                OpLine = MainServer.globalCommond.OperationCode_String,
                LastName = "xlsx",
            }.ToString(), "xlsx");
            ExportExcle(new ExportConfig
            {
                Path1 = exportpath,
                FileName = exportfilename,
            }, MethodeName, rd, null, null, null);
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 結束");
        } // 日 - 票卡進出(完成)

        private void Calculator_EquipAmount()
        {
            var MethodeName = LrtTxnClientSwitch.Day_EquipAmount;
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 結束");
        } // 日 - 設備營收

        private void Calculator_OriginDestination()
        {
            Status = StatusEnum.Calculating;
            MyCommond.WriteLog(ThisReceive, $"初始化變數");
            CardSN = $"{MainServer.OperationDate.ToString("yyyyMMdd")}";
            List<int> ls_O = new List<int>();
            string ls_O_str = "";
            string ReportPath = MyCommond.Path_Program + MyCommond.Path_Report + @"環狀線\" + string.Format(MyCommond.Path_Date, MainServer.OperationDate.ToString("yyyy.MM.dd"));
            MyCommond.WriteLog(ThisReceive, $"檔案路徑2:{ReportPath}");
            var station_Od = MainServer.globalCommond.Stock_Mrt_Y_OdList.ToList();
            var station1 = MainServer.globalCommond.Stock_Mrt_Y_Station;

            List<ThreadList> List_Thread = new List<ThreadList>();

            try
            {
                #region 大OD表
                //GetBidOD:
                MyCommond.WriteLog(ThisReceive, $"大OD表輸出");
                int stationCounts_A = MainServer.globalCommond.Stock_Mrt_Y_Station.Count;
                string[,] ODReport_A = new string[stationCounts_A + 2, stationCounts_A + 2];
                for (int i = 0; i < stationCounts_A; i++)
                {
                    ODReport_A[0, i + 1] = MainServer.globalCommond.Stock_Mrt_Y_Station[i].StationName;
                    ODReport_A[i + 1, 0] = MainServer.globalCommond.Stock_Mrt_Y_Station[i].StationName;
                }
                ODReport_A[0, ODReport_A.GetLength(1) - 1] = "NULL";
                ODReport_A[ODReport_A.GetLength(0) - 1, 0] = "NULL";

                int RideCounts_A = 0;

                ThreadList threadList_BigOd = new ThreadList()
                {
                    Name = "大OD表_一般",
                    thread = new System.Threading.Thread(() =>
                    {
                        foreach (var item in MainServer.Data_Merge0)
                        {
                            if (false) MyCommond.WriteLog(ThisReceive, item.ToString());
                            int EntryIndex = ODReport_A.GetLength(0) - 1;
                            int ExitIndex = ODReport_A.GetLength(1) - 1;
                            if (item.ENTRY_LOC_ID != "NULL") EntryIndex = MainServer.globalCommond.Stock_Mrt_Y_Station.FindIndex(x => x.StationName == item.ENTRY_LOC_ID) + 1;
                            if (item.SVCE_LOC_ID != "NULL") ExitIndex = MainServer.globalCommond.Stock_Mrt_Y_Station.FindIndex(x => x.StationName == item.SVCE_LOC_ID) + 1;
                            if (EntryIndex == 0 || ExitIndex == 0)
                            {
                                MyCommond.WriteLog(ThisReceive, $"Error! EntryIndex:{EntryIndex,4}, ExitIndex:{ExitIndex,4}");
                                continue;
                            }

                            ODReport_A[ExitIndex, EntryIndex] = (Convert.ToInt32(ODReport_A[ExitIndex, EntryIndex]) + 1).ToString();

                            var IsRide = MainServer.globalCommond.Stock_Mrt_Y_OdList.FindIndex(x => (x.Entry == item.ENTRY_LOC_ID || item.ENTRY_LOC_ID == "NULL") && x.Exit == item.SVCE_LOC_ID && x.IsPass == "Y");
                            if (IsRide > -1) RideCounts_A++;
                        }

                        MyCommond.ExportData_BigOD(ThisReceive, new ExportConfig()
                        {
                            Path1 = ReportPath,
                            DTime = MainServer.OperationDate,
                            FirstName = "",
                            FileName = "OD表(大)",
                            SubName = "",
                            OpLine = "環狀線",
                            LastName = "xlsx",
                        }, ODReport_A, RideCounts_A.ToString(), Encoding.UTF8);

                        List_Thread.Find(x => x.Name == "大OD表_一般").IsRun = false;
                    }),
                    IsRun = true,
                };

                List_Thread.Add(threadList_BigOd);

                #endregion

                #region 小OD表_一般
            //GetSmallOD:
                MyCommond.WriteLog(ThisReceive, $"小OD表輸出");
                var stationCounts = MainServer.globalCommond.Stock_Mrt_Y_Station.FindAll(x => x.Company == "新北捷運");
                MyCommond.WriteLog(ThisReceive, $"取得新北捷運車站數量:{stationCounts.Count}");
                string[,] ODReport = new string[stationCounts.Count + 2, stationCounts.Count + 2];
                for (int i = 0; i < stationCounts.Count; i++)
                {
                    ODReport[0, i + 1] = stationCounts[i].StationName;
                    ODReport[i + 1, 0] = stationCounts[i].StationName;
                }
                ODReport[0, ODReport.GetLength(1) - 1] = "NULL";
                ODReport[ODReport.GetLength(0) - 1, 0] = "NULL";

                int RideCounts = 0;

                string[] fdsf = new string[] { "0x1012", "0x1022", "0x1111", "0x1121" /*, "0x1002", "0x1102", "0x1112", "0x1122"*/ };
                var data_merge0_1 = (from p in MainServer.Data_Merge0 where Int32.Parse(p.Remark.Substring(2, p.Remark.Length - 2), System.Globalization.NumberStyles.HexNumber) > 255 && (fdsf.Contains(p.Remark)) select p).ToList();

                ThreadList threadList_SmallOd_Normal = new ThreadList()
                {
                    Name = "小OD表_一般",
                    thread = new System.Threading.Thread(() =>
                    {
                        foreach (var item in data_merge0_1)
                        {
                            if (false) MyCommond.WriteLog(ThisReceive, item.ToString());
                            int EntryIndex = ODReport.GetLength(0) - 1;
                            int ExitIndex = ODReport.GetLength(1) - 1;
                            if (item.M_Entry != "NULL") EntryIndex = stationCounts.FindIndex(x => x.StationName == item.M_Entry) + 1;
                            if (item.M_Exit != "NULL") ExitIndex = stationCounts.FindIndex(x => x.StationName == item.M_Exit) + 1;
                            if (ExitIndex == 0) ExitIndex = stationCounts.FindIndex(x => x.StationName == $"Y{item.M_Exit}") + 1;
                            if (ExitIndex == 0) ExitIndex = stationCounts.FindIndex(x => x.StationName == $"{item.M_Exit}民生") + 1;
                            if (EntryIndex == 0 || ExitIndex == 0)
                            {
                                MyCommond.WriteLog(ThisReceive, $"Error! EntryIndex:{EntryIndex,4}, ExitIndex:{ExitIndex,4}");
                                continue;
                            }

                            ODReport[ExitIndex, EntryIndex] = (Convert.ToInt32(ODReport[ExitIndex, EntryIndex]) + 1).ToString();

                            var IsRide = MainServer.globalCommond.Stock_Mrt_Y_OdList.FindIndex(x => (x.Entry == item.ENTRY_LOC_ID || item.M_Entry == "NULL") && x.Exit == item.M_Exit && x.IsPass == "Y");
                            if (IsRide > -1) RideCounts++;
                        }

                        MyCommond.ExportData_SmallOD(ThisReceive, new ExportConfig()
                        {
                            Path1 = ReportPath,
                            DTime = MainServer.OperationDate,
                            FirstName = "",
                            FileName = "OD表(小)",
                            SubName = "一般",
                            OpLine = "環狀線",
                            LastName = "xlsx",
                        }, ODReport, "0", Encoding.UTF8);

                        List_Thread.Find(x => x.Name == "小OD表_一般").IsRun = false;
                    }),
                    IsRun = true,
                };

                List_Thread.Add(threadList_SmallOd_Normal);

                #endregion

                #region 小OD表_企劃

                string[,] ODReport_Planning = new string[stationCounts.Count + 2, stationCounts.Count + 2];
                for (int i = 0; i < stationCounts.Count; i++)
                {
                    ODReport_Planning[0, i + 1] = stationCounts[i].StationName;
                    ODReport_Planning[i + 1, 0] = stationCounts[i].StationName;
                }
                ODReport_Planning[0, ODReport_Planning.GetLength(1) - 1] = "NULL";
                ODReport_Planning[ODReport_Planning.GetLength(0) - 1, 0] = "NULL";

                ThreadList threadList_SmallOd_Planning = new ThreadList()
                {
                    Name = "小OD表_企劃",
                    thread = new System.Threading.Thread(() =>
                    {
                        foreach (var item in MainServer.Data_Merge0)
                        {
                            if (false) MyCommond.WriteLog(ThisReceive, item.ToString());
                            int EntryIndex = ODReport_Planning.GetLength(0) - 1;
                            int ExitIndex = ODReport_Planning.GetLength(1) - 1;
                            if (item.M_Entry != "NULL") EntryIndex = stationCounts.FindIndex(x => x.StationName == item.M_Entry) + 1;
                            if (item.M_Exit != "NULL") ExitIndex = stationCounts.FindIndex(x => x.StationName == item.M_Exit) + 1;
                            if (ExitIndex == 0) ExitIndex = stationCounts.FindIndex(x => x.StationName == $"Y{item.M_Exit}") + 1;
                            if (ExitIndex == 0) ExitIndex = stationCounts.FindIndex(x => x.StationName == $"{item.M_Exit}民生") + 1;
                            if (EntryIndex == 0 || ExitIndex == 0)
                            {
                                MyCommond.WriteLog(ThisReceive, $"Error! EntryIndex:{EntryIndex,4}, ExitIndex:{ExitIndex,4}");
                                continue;
                            }

                            ODReport_Planning[ExitIndex, EntryIndex] = (Convert.ToInt32(ODReport_Planning[ExitIndex, EntryIndex]) + 1).ToString();

                            var IsRide = MainServer.globalCommond.Stock_Mrt_Y_OdList.FindIndex(x => (x.Entry == item.ENTRY_LOC_ID || item.M_Entry == "NULL") && x.Exit == item.M_Exit && x.IsPass == "Y");
                            if (IsRide > -1) RideCounts++;
                        }

                        MyCommond.ExportData_SmallOD(ThisReceive, new ExportConfig()
                        {
                            Path1 = ReportPath,
                            DTime = MainServer.OperationDate,
                            FirstName = "",
                            FileName = "OD表(小)",
                            SubName = "企劃",
                            OpLine = "環狀線",
                            LastName = "xlsx",
                        }, ODReport_Planning, "0", Encoding.UTF8);

                        List_Thread.Find(x => x.Name == "小OD表_企劃").IsRun = false;
                    }),
                    IsRun = true,
                };

                List_Thread.Add(threadList_SmallOd_Planning);

                #endregion
            }
            catch (Exception ex)
            {
                throw ex;
            }

            ThreadRun(List_Thread);

        FinishSub:
            MyCommond.WriteLog(ThisReceive, $"結束本副程式輸出");
            Status = StatusEnum.Finish;
        } // 日 - 起訖總表(完成)

        private void Calculator_TrafficAmount()
        {
            var MethodeName = LrtTxnClientSwitch.Day_TrafficAmount;
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 結束");
        } // 日 - 營收報表

        #endregion

        #region 月

        private void Calculator_Month_AllRideList()
        {
            var MethodeName = LrtTxnClientSwitch.Month_AllRideList;
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");
        } // 月 - 

        private void Calculator_Month_OriginDestination()
        {
            var MethodeName = LrtTxnClientSwitch.Month_OriginDestination;
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");
        } // 月 - 起訖總表

        private void Calculator_Month_ElectronicTicket_Station()
        {
            var MethodeName = LrtTxnClientSwitch.Month_ElectronicTicket_Station;
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");
        } // 月 - 票卡進出(車站)

        private void Calculator_Month_ElectronicTicket_Day()
        {
            var MethodeName = LrtTxnClientSwitch.Month_ElectronicTicket_Day;
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");
        } // 月 - 票卡進出(天期)

        private void Calculator_Month_OwnTicketVolume_Station()
        {
            var MethodeName = LrtTxnClientSwitch.Month_OwnTicketVolume_Station;
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");
        } // 月 - 自有票卡(車站)

        private void Calculator_Month_OwnTicketVolume_Day()
        {
            var MethodeName = LrtTxnClientSwitch.Month_OwnTicketVolume_Day;
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");
        } // 月 - 自有票卡(天期)

        private void Calculator_Month_TrafficAmount_Station()
        {
            var MethodeName = LrtTxnClientSwitch.Month_TrafficAmount_Station;
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");
        } // 月 - 營收報表(車站)

        private void Calculator_Month_TrafficAmount_Day()
        {
            var MethodeName = LrtTxnClientSwitch.Month_TrafficAmount_Day;
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");
        } // 月 - 營收報表(天期)

        private void Calculator_Month_EquipAmount_Station()
        {
            var MethodeName = LrtTxnClientSwitch.Month_EquipAmount_Station;
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");
        } // 月 - 設備營收(車站)

        private void Calculator_Month_EquipAmount_Day()
        {
            var MethodeName = LrtTxnClientSwitch.Month_EquipAmount_Day;
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");
        } // 月 - 設備營收(天期)

        private void Calculator_Month_Volumn()
        {
            var MethodeName = LrtTxnClientSwitch.Month_Volume;
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");
        } // 月 - 每日運量總表(運量)

        private void Calculator_Month_Amount()
        {
            var MethodeName = LrtTxnClientSwitch.Month_Amount;
            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            MyCommond.WriteLog(ThisReceive, $"{MethodeName.ToDescription()} - 開始");
        } // 月 - 每日運量總表(營收)

        #endregion

        #endregion

        #region 加值報表分析(未完成)

        private void Calculator_DayAddValueAnalyze()
        {
            Status = StatusEnum.Calculating;
            CardSN = $"{MainServer.OperationDate.ToString("yyyyMMdd")}";

            // 篩選TxnType是加值相關的
            string[] Filter_AddValue = (from p in MainServer.globalCommond.Stock_Mrt_TxnType 
                                        where p.mCheck == "AA" 
                                        select p.Num.ToString()).ToArray();
            string[] Filter_AddValue_Cancal = (from p in MainServer.globalCommond.Stock_Mrt_TxnType 
                                               where p.mCheck == "AF" 
                                               select p.Num.ToString()).ToArray();

            // 篩選TxnSubType為各類加值的
            // 自動加值-車站 => CARD_TXN_SUBTYPE_ID:20(閘門自動加值)、64(PAM_ATIM自動加值)
            string[] Filter_SubType_Auto = (from p in MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ecc 
                                            where p.Num == 20 || p.Num == 64 
                                            select p.Num.ToString()).ToArray();

            // 人工加值-車站 => CARD_TXN_SUBTYPE_ID:17(現金)、22(溢扣處理)、106(捷運溢扣)、139(加值回饋金)、181(自動售票機加值)、140(累計金額為3萬元以上)
            //                  DEV_ID:207(PAM)
            string[] Filter_SubType_Artificial = (from p in MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ecc 
                                                  where p.Num == 17 || p.Num == 22 || p.Num == 106 || p.Num == 139 || p.Num == 181 || p.Num == 140 
                                                  select p.Num.ToString()).ToArray();

            // 機具加值-車站 => CARD_TXN_SUBTYPE_ID:17(現金)、22(溢扣處理)、106(捷運溢扣)、139(加值回饋金)、181(自動售票機加值)
            //                  DEV_ID:206(GATE)、208(ATIM)
            string[] Filter_SubType_Machine = (from p in MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ecc 
                                               where p.Num == 17 || p.Num == 22 || p.Num == 106 || p.Num == 139 || p.Num == 181 
                                               select p.Num.ToString()).ToArray();

            // 篩選設備
            int[] Filter_Device_Artificial = new int[] { 207 };
            int[] Filter_Device_Machine = new int[] { 206, 208 };

            // 篩選公司車站
            List<MrtStationList> Filter_Station = MainServer.globalCommond.Stock_Mrt_Y_Station.FindAll(x => x.Company == "新北捷運").ToList();

            // 加值欄位位置
            List<String_String_Int_Int_Int> Report_Column_Temp = (from p in MainServer.globalCommond.Stock_Mrt_IssureCode 
                                                                  where p.AddValueTag != 0 
                                                                  select p).ToList();

            var Filter_Data_1 = (from p in MainServer.TxnAllData 
                                 where Filter_AddValue.Contains(p.CARD_TXN_TYPE_ID) 
                                 select p).ToList();
            MyCommond.WriteLog(ThisReceive, $"篩選出加值筆數:{Filter_Data_1.Count}");

            var Filter_Data_2 = (from p in MainServer.TxnAllData 
                                 where Filter_AddValue_Cancal.Contains(p.CARD_TXN_TYPE_ID) 
                                 select p).ToList();
            MyCommond.WriteLog(ThisReceive, $"篩選出取消加值筆數:{Filter_Data_2.Count}");

            double[,] rd = new double[Filter_Station.Count, Report_Column_Temp.Count * 2 * 3];

            try
            {
                // 自動加值 offset = 0
                List<TxnTransaction> AddValue_Auto = (from p in Filter_Data_1 
                                                      where Filter_SubType_Auto.Contains(p.CARD_TXN_SUBTYPE_ID) 
                                                      select p).ToList();
                MyCommond.WriteLog(ThisReceive, $"\t自動加值筆數:{AddValue_Auto.Count}");
                CompareAndAddData(AddValue_Auto, /*-------*/ Filter_Station, Report_Column_Temp, 0, true, ref rd);

                // 人工加值 offset = 1
                List<TxnTransaction> AddValue_Artificial = (from p in Filter_Data_1 
                                                            where Filter_SubType_Artificial.Contains(p.CARD_TXN_SUBTYPE_ID) && Filter_Device_Artificial.Contains(Convert.ToInt16(p.DEV_ID.Substring(0,3))) 
                                                            select p).ToList();
                MyCommond.WriteLog(ThisReceive, $"\t人工加值筆數:{AddValue_Artificial.Count}");
                CompareAndAddData(AddValue_Artificial, /*-*/ Filter_Station, Report_Column_Temp, 1, true, ref rd);
                List<TxnTransaction> AddValue_Cancal = (from p in Filter_Data_1 
                                                        where Filter_AddValue_Cancal.Contains(p.CARD_TXN_TYPE_ID) 
                                                        select p).ToList();
                MyCommond.WriteLog(ThisReceive, $"\t取消加值筆數:{AddValue_Cancal.Count}");
                CompareAndAddData(Filter_Data_2, /*-------*/ Filter_Station, Report_Column_Temp, 1, false, ref rd);

                // 機具加值 offset = 2
                List<TxnTransaction> AddValue_Machine = (from p in Filter_Data_1 
                                                         where Filter_SubType_Machine.Contains(p.CARD_TXN_SUBTYPE_ID) && Filter_Device_Machine.Contains(Convert.ToInt16(p.DEV_ID.Substring(0,3))) 
                                                         select p).ToList();
                MyCommond.WriteLog(ThisReceive, $"\t機具加值筆數:{AddValue_Machine.Count}");
                CompareAndAddData(AddValue_Machine, /*----*/ Filter_Station, Report_Column_Temp, 2, true, ref rd);

                // 輸出
                string exportpath = MyCommond.Path_Program + MyCommond.Path_Report + @"環狀線\" + string.Format(MyCommond.Path_Date, MainServer.OperationDate.ToString("yyyy.MM.dd"));
                ReportFileTemp2 cheeee = MainServer.globalCommond.Stock_ReportFile.Find(x => 
                                            x.TransportSystem == MainServer.globalCommond.UsingSystem && 
                                            x.OperationLine == MainServer.globalCommond.OperationCode && 
                                            x.ReportNameEng == MrtTxnClientSwitch.Day_AddValue.ToString());
                string temp_Filename = cheeee.ReportNameCh;

                // 比對用資料
                bool compareDataExport = false;
                if (compareDataExport) MyCommond.ExportData<TxnTransaction>(ThisReceive, new ExportConfig()
                {
                    Path1 = exportpath,
                    DTime = MainServer.OperationDate,
                    yearType = YearType.RC,
                    FirstName = "",
                    FileName = $"加值明細",
                    SubName = "",
                    OpLine = MainServer.globalCommond.OperationCode_String,
                    LastName = "xlsx",
                }, new TxnTransaction().ExportTitle_zhTW(), Filter_Data_1, Encoding.UTF8);
                if (compareDataExport) MyCommond.ExportData<TxnTransaction>(ThisReceive, new ExportConfig()
                {
                    Path1 = exportpath,
                    DTime = MainServer.OperationDate,
                    yearType = YearType.RC,
                    FirstName = "",
                    FileName = $"取消加值明細",
                    SubName = "",
                    OpLine = MainServer.globalCommond.OperationCode_String,
                    LastName = "xlsx",
                }, new TxnTransaction().ExportTitle_zhTW(), Filter_Data_2, Encoding.UTF8);

                // 資料輸出
                string exportfilename = MyCommond.CheckAndReName(exportpath, new ExportConfig()
                {
                    Path1 = exportpath,
                    DTime = MainServer.OperationDate,
                    yearType = YearType.RC,
                    FirstName = "",
                    FileName = $"{temp_Filename}",
                    SubName = "",
                    OpLine = MainServer.globalCommond.OperationCode_String,
                    LastName = "xlsx",
                }.ToString(), "xlsx");
                ExportExcle(new ExportConfig
                {
                    Path1 = exportpath,
                    FileName = exportfilename,
                }, MrtTxnClientSwitch.Day_AddValue, rd, null, null, null);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            Status = StatusEnum.Finish;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="listData">交易資料</param>
        /// <param name="listRow">車站(行)</param>
        /// <param name="listColumn">票證公司(欄)</param>
        /// <param name="offset">各種加值(位移)</param>
        /// <param name="result">報表欄位陣列</param>
        private void CompareAndAddData(List<TxnTransaction> listData, List<MrtStationList> listRow, List<String_String_Int_Int_Int> listColumn, int offset, bool add, ref double[,] result)
        {
            foreach(TxnTransaction data in listData)
            {
                try
                {
                    int mRow = listRow.FindIndex(x => x.CodeName == data.SVCE_LOC_ID);
                    int mCol = listColumn.Find(x => x.Num.ToString() == data.ISSUER_ID).AddValueTag - 1;

                    bool tempCheck = true;
                    if (tempCheck) MyCommond.WriteLog(ThisReceive, $"交易車站({mRow,4}):{data.SVCE_LOC_ID,4}, 票證公司({mCol,4}):{data.ISSUER_ID,4}, Offset:{offset,4}, 交易金額:{data.TXN_AMT,6}元");

                    if (mRow >= 0)
                    {
                        if (add)
                        {
                            result[mRow, (mCol * 6) + (offset * 2) + 0]++;
                            result[mRow, (mCol * 6) + (offset * 2) + 1] += Convert.ToInt16(data.TXN_AMT);
                        }
                        else
                        {
                            result[mRow, (mCol * 6) + (offset * 2) + 0]--;
                            result[mRow, (mCol * 6) + (offset * 2) + 1] -= Convert.ToInt16(data.TXN_AMT);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"Error!Data:{data.ToString()}");
                }
            }

            int a1 = result.GetLength(0);
            int a2 = result.GetLength(1);

            for (int i = 0; i < result.GetLength(0); i++)
            {
                string ss = "";
                for (int j = 0; j < result.GetLength(1); j++)
                {
                    double d1 = result[i, j];
                    ss += $"{d1,-6},";
                }
                MyCommond.WriteLog(ThisReceive, $"result({i,2}) => {ss}");
            }
        }

        #endregion

        #region 各家支付業者運量日報(1/10完成，1/12測試通過)

        private void Calculator_EveryIssuerAnalyze()
        {
            Status = StatusEnum.Calculating;
            CardSN = MainServer.OperationDate.ToString("yyyyMMdd");

            try
            {
                var stationCounts = MainServer.globalCommond.Stock_Mrt_Y_Station.FindAll(x => x.Company == "新北捷運").ToList();
                var ColumnString = MainServer.globalCommond.Stock_Mrt_IssureCode.FindAll(x => x.IssuerRawTag != 0);
                double[,] rd = new double[stationCounts.Count, ColumnString.Count * 2];

                string[] fdsf = new string[] { "0x1012", "0x1022", "0x1111", "0x1121" /*, "0x1002", "0x1102", "0x1112", "0x1122"*/ };
                // 轉換 Remark 16進位變10進位 後大於 255 且 符合上述4個篩選
                var data_merge0_1 = (from p in MainServer.Data_Merge0 where Int32.Parse(p.Remark.Substring(2, p.Remark.Length - 2), System.Globalization.NumberStyles.HexNumber) > 255 && (fdsf.Contains(p.Remark)) select p).ToList();
                try
                {
                    foreach (var item in data_merge0_1)
                    {
                        int Offset_Column = -1;
                        try
                        {
                            Offset_Column = ColumnString.Find(x => x.ChName == item.ISSUER_ID).IssuerRawTag;
                            Offset_Column--;
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"ERROR\nOffset_Column:{Offset_Column}\nitem:{item.ToString()}");
                        }

                        int Offset_Row_Entry = -1;
                        try
                        {
                            Offset_Row_Entry = StationRowCheck(stationCounts, item.M_Entry);
                            //if (item.M_Entry != "NULL") { Offset_Row_Entry = stationCounts.FindIndex(x => x.StationName == item.M_Entry); }
                            //if (Offset_Row_Entry == -1) { Offset_Row_Entry = stationCounts.FindIndex(x => x.StationName == $"Y{item.M_Entry}"); }
                            //if (Offset_Row_Entry == -1) { Offset_Row_Entry = stationCounts.FindIndex(x => x.StationName == $"{item.M_Entry}民生"); }
                            //if (Offset_Row_Entry != -1) { }
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"ERROR\nOffset_Row_Entry:{Offset_Row_Entry}\nitem:{item.ToString()}");
                        }

                        int Offset_Row_Exit = -1;
                        try
                        {
                            Offset_Row_Exit = StationRowCheck(stationCounts, item.M_Exit);
                            //if (item.M_Exit != "NULL") { Offset_Row_Exit = stationCounts.FindIndex(x => x.StationName == item.M_Exit); }
                            //if (Offset_Row_Exit == -1) { Offset_Row_Exit = stationCounts.FindIndex(x => x.StationName == $"Y{item.M_Exit}"); }
                            //if (Offset_Row_Exit == -1) { Offset_Row_Exit = stationCounts.FindIndex(x => x.StationName == $"{item.M_Exit}民生"); }
                            //if (Offset_Row_Exit != -1) { }
                        }
                        catch (Exception ex)
                        {
                            MyCommond.WriteLog(ThisReceive, $"ERROR\nOffset_Row_Exit:{Offset_Row_Exit}\nitem:{item.ToString()}");
                        }

                        if (Offset_Column > -1 && Offset_Row_Entry > -1 && Offset_Row_Exit > -1)
                        {
                            rd[Offset_Row_Entry, Offset_Column * 2 + 0]++;
                            rd[Offset_Row_Exit, Offset_Column * 2 + 1]++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                string exportpath = MyCommond.Path_Program + MyCommond.Path_Report + @"環狀線\" + string.Format(MyCommond.Path_Date, MainServer.OperationDate.ToString("yyyy.MM.dd"));

                string temp_TransportSystem = MainServer.globalCommond.UsingSystem;
                string temp_OperationLine = MainServer.globalCommond.OperationCode;
                string temp_ReportNameEnglish = MrtTxnClientSwitch.Day_EveryIssuer.ToString();

                ReportFileTemp2 cheeee = MainServer.globalCommond.Stock_ReportFile.Find(x => x.TransportSystem == temp_TransportSystem && x.OperationLine == temp_OperationLine && x.ReportNameEng == temp_ReportNameEnglish);
                string temp_Filename = cheeee.ReportNameCh;
                string exportfilename = MyCommond.CheckAndReName(exportpath, new ExportConfig()
                {
                    Path1 = exportpath,
                    DTime = MainServer.OperationDate,
                    yearType = YearType.RC,
                    FirstName = "",
                    FileName = $"{temp_Filename}",
                    SubName = "",
                    OpLine = MainServer.globalCommond.OperationCode_String,
                    LastName = "xlsx",
                }.ToString(), "xlsx");
                ExportExcle(new ExportConfig
                {
                    Path1 = exportpath,
                    FileName = exportfilename,
                }, MrtTxnClientSwitch.Day_EveryIssuer, rd, null, null, null);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            Status = StatusEnum.Finish;
        }
        #endregion

        #endregion

        #region 未使用

        #region C10

        private void Calculator_C10Analyze()
        {
            Status = StatusEnum.Calculating;
            CardSN = MainServer.OperationDate.ToString("yyyyMMdd");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            Status = StatusEnum.Finish;
        }
        #endregion

        #region C11

        private void Calculator_C11Analyze()
        {
            Status = StatusEnum.Calculating;
            CardSN = MainServer.OperationDate.ToString("yyyyMMdd");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            Status = StatusEnum.Finish;
        }
        #endregion

        #region C12

        private void Calculator_C12Analyze()
        {
            Status = StatusEnum.Calculating;
            CardSN = MainServer.OperationDate.ToString("yyyyMMdd");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            Status = StatusEnum.Finish;
        }
        #endregion

        #region C13

        private void Calculator_C13Analyze()
        {
            Status = StatusEnum.Calculating;
            CardSN = MainServer.OperationDate.ToString("yyyyMMdd");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            Status = StatusEnum.Finish;
        }
        #endregion

        #region C14

        private void Calculator_C14Analyze()
        {
            Status = StatusEnum.Calculating;
            CardSN = MainServer.OperationDate.ToString("yyyyMMdd");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            Status = StatusEnum.Finish;
        }
        #endregion

        #region C15

        private void Calculator_C15Analyze()
        {
            Status = StatusEnum.Calculating;
            CardSN = MainServer.OperationDate.ToString("yyyyMMdd");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            Status = StatusEnum.Finish;
        }
        #endregion

        #region C16

        private void Calculator_C16Analyze()
        {
            Status = StatusEnum.Calculating;
            CardSN = MainServer.OperationDate.ToString("yyyyMMdd");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            Status = StatusEnum.Finish;
        }
        #endregion

        #region C17

        private void Calculator_C17Analyze()
        {
            Status = StatusEnum.Calculating;
            CardSN = MainServer.OperationDate.ToString("yyyyMMdd");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            Status = StatusEnum.Finish;
        }
        #endregion

        #region C18

        private void Calculator_C18Analyze()
        {
            Status = StatusEnum.Calculating;
            CardSN = MainServer.OperationDate.ToString("yyyyMMdd");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            Status = StatusEnum.Finish;
        }
        #endregion

        #region C19

        private void Calculator_C19Analyze()
        {
            Status = StatusEnum.Calculating;
            CardSN = MainServer.OperationDate.ToString("yyyyMMdd");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            Status = StatusEnum.Finish;
        }
        #endregion

        #region C20

        private void Calculator_C20Analyze()
        {
            Status = StatusEnum.Calculating;
            CardSN = MainServer.OperationDate.ToString("yyyyMMdd");

            try
            {

            }
            catch (Exception ex)
            {
                throw ex;
            }

            Status = StatusEnum.Finish;
        }
        #endregion

        #endregion

        #region 交易資料重新編寫

        private string Check_TxnType(string m_Data)
        {
            string Result_str = m_Data;
            List<String_String_String_String_Int> Result_List = MainServer.globalCommond.Stock_Mrt_TxnType;
            if (Result_List.FindIndex(x => x.Num == Convert.ToInt32(m_Data)) > -1)
            {
                Result_str = Result_List.Find(x => x.Num == Convert.ToInt32(m_Data)).ChName;
            }
            return Result_str;
        }

        private string Check_TxnSubType(Ele_Iss Is_Id, string m_Data)
        {
            string Result_str = m_Data;
            List<String_String_String_String_Int> Result_List = new List<String_String_String_String_Int>();

            if (Is_Id == Ele_Iss.ECC) /*------------*/ { Result_List = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ecc; }
            else if (Is_Id == Ele_Iss.Ipass) /*-----*/ { Result_List = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ipass; }
            else if (Is_Id == Ele_Iss.Icash) /*-----*/ { Result_List = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Icash; }
            else if (Is_Id == Ele_Iss.OwnTicket) /*-*/ { Result_List = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ecc; }
            else /*---------------------------------*/ { return Result_str; }

            if (Result_List.FindIndex(x => x.mCheck == "o" && x.Num == Convert.ToInt32(m_Data)) > -1)
            {
                Result_str = Result_List.Find(x => x.mCheck == "o" && x.Num == Convert.ToInt32(m_Data)).ChName;
            }
            return Result_str;
        }

        private string Check_IssuerId(string m_Data)
        {
            string Result_str = m_Data;
            List<String_String_Int_Int_Int> Result_List = MainServer.globalCommond.Stock_Mrt_IssureCode;
            if (Result_List.FindIndex(x => x.Num == Convert.ToInt32(m_Data)) > -1)
            {
                Result_str = Result_List.Find(x => x.Num == Convert.ToInt32(m_Data)).ChName.Split('/')[0];
            }
            return Result_str;
        }

        private string Check_Station(string m_Data)
        {
            string Result_str = m_Data;
            List<MrtStationList> Result_List = MainServer.globalCommond.Stock_Mrt_Y_Station;
            if (Result_List.FindIndex(x => x.CodeName == m_Data) > -1)
            {
                Result_str = Result_List.Find(x => x.CodeName == m_Data).StationName;
            }
            return Result_str;
        }

        private string Check_Fare(string m_Data)
        {
            string Result_str = m_Data;
            List<String_String_String_Int> Result_List = MainServer.globalCommond.Stock_Mrt_FareProductType;
            if (Result_List.FindIndex(x => x.Num == Convert.ToInt32(m_Data)) > -1)
            {
                Result_str = Result_List.Find(x => x.Num == Convert.ToInt32(m_Data)).ChName;
            }
            return Result_str;
        }

        private string Check_Ticket(string m_Data)
        {
            string Result_str = m_Data;
            List<String_String_String_String_Int> Result_List = MainServer.globalCommond.Stock_Mrt_TxnTypeSub_Ticket;
            if (Result_List.FindIndex(x => x.Num == Convert.ToInt32(m_Data)) > -1)
            {
                Result_str = Result_List.Find(x => x.Num == Convert.ToInt32(m_Data)).ChName;
            }
            return Result_str;
        }

        private string Check_IO(List<MrtStationOdList> results, E2 IO_, string m_Entry, string m_Exit)
        {
            try
            {
                var Find_Data = results.Find(x => x.Entry == m_Entry && x.Exit == m_Exit);
                if (E2.Entry == IO_) { return Find_Data.Entry_Own; }
                else if (E2.Exit == IO_) { return Find_Data.Exit_Own; }
                else { return "E1"; }
            }
            catch (Exception e)
            {
                if (E2.Entry == IO_) { return m_Entry; }
                else if (E2.Exit == IO_) { return m_Exit; }
                else { return "E2"; }
            }
        }

        private string Check_TransferMark(List<string> C_Data, string m_Entry, string m_Exit)
        {
            try
            {
                if (C_Data.Contains(m_Entry) || C_Data.Contains(m_Exit)) { return "@"; }
                return "";
            }
            catch (Exception e)
            {
                return "";
            }
            return "";
        }

        private string Check_Remark(string m_Entry, string m_Exit)
        {
            string a1 = "", a2 = "";
            int s1 = 12 - MyCommond.StringCount(m_Entry);
            int s2 = 12 - MyCommond.StringCount(m_Exit);
            for (int i = 0; i < s1; i++) { a1 += " "; }
            for (int i = 0; i < s2; i++) { a2 += " "; }
            string ErrorString = $"\nm_Entry:{a1}{m_Entry}, m_Exit:{a2}{m_Exit}\n";
            string[] asdf = new string[] { "大坪林", "景安", "頭前庄" };

            try
            {
                string re = "";
                var RouteList = MainServer.globalCommond.Stock_Mrt_Y_OdList.ToList();
                var station = MainServer.globalCommond.Stock_Mrt_Y_Station.ToList();

                var Route = RouteList.Find(x => x.Entry == m_Entry && x.Exit == m_Exit);

                string Byte1 = (((Route != null) ? ((Route.IsPass == "Y") ? "1" : "0") : "0"));
                string Byte2 = ((asdf.Contains(m_Exit)) ? "1" : "0");
                string EntryCompany = "0";
                string ExitCompany = "0";

                ErrorString += $"初始化變數\t";

                try
                {
                    ErrorString += $"判斷進站\t";

                    string tempI = station.Find(x => x.StationName == m_Entry).Company;
                    if (tempI == "台北捷運") { EntryCompany = "1"; }
                    else if (tempI == "新北捷運") { EntryCompany = "2"; }
                }
                catch(Exception e)
                {
                    ErrorString += $"\n進站 - ERROR!\n{e.Message}\n";
                }

                try
                {
                    ErrorString += $"判斷出站\n";

                    string tempO = station.Find(x => x.StationName == m_Exit).Company;
                    if (tempO == "台北捷運") { ExitCompany = "1"; }
                    else if (tempO == "新北捷運") { ExitCompany = "2"; }
                }
                catch (Exception e)
                {
                    ErrorString += $"\n出站 - ERROR!\n{e.Message}\n";
                }

                ErrorString += $"編寫代號 => Route.IsPass:{((Route != null) ? Route.IsPass:"E"),4}, EntryCompany:{EntryCompany,4}, ExitCompany:{ExitCompany,4}\n";
                re += $"0x{Byte1}{Byte2}{EntryCompany}{ExitCompany}";
                ErrorString += $"輸出:{re}";
                return re;
            }
            catch (Exception e)
            {
                return "";
            }
            finally
            {
                Console.WriteLine(ErrorString);
            }
        }

        #endregion

        private int StationRowCheck(List<MrtStationList> msl, string StationName)
        {
            int result = -1;
            if (StationName != "NULL") { result = msl.FindIndex(x => x.StationName == StationName); }
            if (result == -1) { result = msl.FindIndex(x => x.StationName == $"Y{StationName}"); }
            if (result == -1) { result = msl.FindIndex(x => x.StationName == $"{StationName}民生"); }
            if (result != -1) { }
            return result;
        }

        private string[,] OdForPlanning()
        {
            int RideCounts = 0;

            var stationCounts = MainServer.globalCommond.Stock_Mrt_Y_Station.FindAll(x => x.Company == "新北捷運");
            string[,] ODReport_Planning = new string[stationCounts.Count + 2, stationCounts.Count + 2];
            for (int i = 0; i < stationCounts.Count; i++)
            {
                ODReport_Planning[0, i + 1] = stationCounts[i].StationName;
                ODReport_Planning[i + 1, 0] = stationCounts[i].StationName;
            }
            ODReport_Planning[0, ODReport_Planning.GetLength(1) - 1] = "NULL";
            ODReport_Planning[ODReport_Planning.GetLength(0) - 1, 0] = "NULL";

            foreach (var item in MainServer.Data_Merge0)
            {
                if (false) MyCommond.WriteLog(ThisReceive, item.ToString());
                int EntryIndex = ODReport_Planning.GetLength(0) - 1;
                int ExitIndex = ODReport_Planning.GetLength(1) - 1;
                if (item.M_Entry != "NULL") EntryIndex = stationCounts.FindIndex(x => x.StationName == item.M_Entry) + 1;
                if (item.M_Exit != "NULL") ExitIndex = stationCounts.FindIndex(x => x.StationName == item.M_Exit) + 1;
                if (ExitIndex == 0) ExitIndex = stationCounts.FindIndex(x => x.StationName == $"Y{item.M_Exit}") + 1;
                if (ExitIndex == 0) ExitIndex = stationCounts.FindIndex(x => x.StationName == $"{item.M_Exit}民生") + 1;
                if (EntryIndex == 0 || ExitIndex == 0)
                {
                    MyCommond.WriteLog(ThisReceive, $"Error! EntryIndex:{EntryIndex,4}, ExitIndex:{ExitIndex,4}");
                    continue;
                }

                ODReport_Planning[ExitIndex, EntryIndex] = (Convert.ToInt32(ODReport_Planning[ExitIndex, EntryIndex]) + 1).ToString();

                var IsRide = MainServer.globalCommond.Stock_Mrt_Y_OdList.FindIndex(x => (x.Entry == item.ENTRY_LOC_ID || item.M_Entry == "NULL") && x.Exit == item.M_Exit && x.IsPass == "Y");
                if (IsRide > -1) RideCounts++;
            }

            return ODReport_Planning;
        }

        private bool ThreadRun(List<ThreadList> Lt)
        {
            MyCommond.WriteLog(ThisReceive, $"執行線程 - 開始"); foreach (var item in Lt) { Lt.Find(x => x == item).thread.Start(); }
            MyCommond.WriteLog(ThisReceive, $"執行線程 - 等待"); foreach (var item in Lt) { Lt.Find(x => x == item).thread.Join(); }
            MyCommond.WriteLog(ThisReceive, $"執行線程 - 結束");
            return false;
        }

        public bool ExportExcle_Old(ExportConfig ec, LrtTxnClientSwitch MethodeName, int[,] Data_Mrt, int[,] Data_MobilePay, int[,] Data_OwnTicket, int[] Data_Other = null)
        {
            string exportFullPath = $"{ec.Path1}{ec.ToString()}";

            #region 讀取輸出位置資料

            MyCommond.WriteLog(ThisReceive, $"取得輸出資料變數");
            string mYear = "", mMonth = "", mDay = "";

            string template = MyCommond.Path_Program + MyCommond.Path_Template + MethodeName.ToString() + ".xlsx";

            bool ExportCheck = false;
            bool Customized = (MethodeName == LrtTxnClientSwitch.Day_Volume || MethodeName == LrtTxnClientSwitch.Day_Amount/* || MethodeName == LrtTxnClientSwitch.Day_EachStation_EachTime*/);

            ReportFileTemp2 reportFileTemp = new ReportFileTemp2();
            List<ReportFileTemp2> S_R_F = new List<ReportFileTemp2>();
            int S_R_F_Index = 0;
            string MethodeName_zhTW = "";
            string MethodeName_String = "";
            int ReportDateStartRow /*----*/ = 0;
            int ReportDateStartColumn /*-*/ = 0;
            int StartRow /*--------------*/ = 0;
            int StartColumn /*-----------*/ = 0;
            int EndRow /*----------------*/ = 0;
            int EndColumn /*-------------*/ = 0;
            int thisBreakOffRow /*-------*/ = 0;
            int thisBreakOffRowNum /*----*/ = 0;
            int thisBreakOffColumn /*----*/ = 0;
            int thisBreakOffColumnNum /*-*/ = 0;
            int MobilePayColumn /*-------*/ = 0;
            int OwnTiketColumn /*--------*/ = 0;

            if (true) { goto ExportConfigSetting_New; } else { goto ExportConfigSetting_Old; }

        ExportConfigSetting_Old:
            #region 舊版設定
            S_R_F = MainServer.globalCommond.Stock_ReportFile.FindAll(x => x.OperationLine == MainServer.globalCommond.OperationCode).ToList();
            S_R_F_Index = MainServer.globalCommond.Stock_ReportFile.FindIndex(x => x.OperationLine == MainServer.globalCommond.OperationCode);
            MyCommond.WriteLog(ThisReceive, $"找到 {((S_R_F_Index > -1) ? $"{S_R_F.Count}" : "0"),4} 筆資料");
            if (S_R_F_Index == -1) { return false; }


            MethodeName_zhTW = MethodeName.ToDescription();
            MethodeName_String = MethodeName.ToString();

            ReportDateStartRow /*----*/ = Convert.ToInt16((/*----------*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).ReportDateRow == "-") /*-----*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).ReportDateRow)); /*-----*/ MyCommond.WriteLog(ThisReceive, "產製日期起始行：" + ReportDateStartRow);
            ReportDateStartColumn /*-*/ = MyCommond.ExcelColumnNum((/*-*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).ReportDateColumn == "-") /*--*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).ReportDateColumn)); /*--*/ MyCommond.WriteLog(ThisReceive, "產製日期起始欄：" + ReportDateStartColumn);
            StartRow /*--------------*/ = Convert.ToInt16((/*----------*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).StartRow == "-") /*----------*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).StartRow)); /*----------*/ MyCommond.WriteLog(ThisReceive, "輸出資料起始行：" + StartRow);
            StartColumn /*-----------*/ = MyCommond.ExcelColumnNum((/*-*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).StartColumn == "-") /*-------*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).StartColumn)); /*-------*/ MyCommond.WriteLog(ThisReceive, "輸出資料起始欄：" + StartColumn);
            EndRow /*----------------*/ = Convert.ToInt16((/*----------*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).EndRow == "-") /*------------*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).EndRow)); /*------------*/ MyCommond.WriteLog(ThisReceive, "輸出資料結束行：" + EndRow);
            EndColumn /*-------------*/ = MyCommond.ExcelColumnNum((/*-*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).EndColumn == "-") /*---------*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).EndColumn)); /*---------*/ MyCommond.WriteLog(ThisReceive, "輸出資料結束欄：" + EndColumn);
            thisBreakOffRow /*-------*/ = Convert.ToInt16((/*----------*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffRow == "-") /*-------*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffRow)); /*-------*/ MyCommond.WriteLog(ThisReceive, "資料中斷行    ：" + thisBreakOffRow);
            thisBreakOffRowNum /*----*/ = Convert.ToInt16((/*----------*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffRowNum == "-") /*----*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffRowNum)); /*----*/ MyCommond.WriteLog(ThisReceive, "資料中斷行數  ：" + thisBreakOffRowNum);
            thisBreakOffColumn /*----*/ = MyCommond.ExcelColumnNum((/*-*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffColumn == "-") /*----*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffColumn)); /*----*/ MyCommond.WriteLog(ThisReceive, "資料中斷欄    ：" + thisBreakOffColumn);
            thisBreakOffColumnNum /*-*/ = Convert.ToInt16((/*----------*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffColumnNum == "-") /*-*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).BreakOffColumnNum)); /*-*/ MyCommond.WriteLog(ThisReceive, "資料中斷欄數  ：" + thisBreakOffColumnNum);
            MobilePayColumn /*-------*/ = MyCommond.ExcelColumnNum((/*-*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).MobilePayColumn == "-") /*---*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).MobilePayColumn)); /*---*/ MyCommond.WriteLog(ThisReceive, "電子支付起始欄：" + MobilePayColumn);
            OwnTiketColumn /*--------*/ = MyCommond.ExcelColumnNum((/*-*/ (S_R_F.Find(x => x.ReportNameEng == MethodeName_String).OwnTiketColumn == "-") /*----*/ ? "0" : S_R_F.Find(x => x.ReportNameEng == MethodeName_String).OwnTiketColumn)); /*----*/ MyCommond.WriteLog(ThisReceive, "自有票起始欄  ：" + OwnTiketColumn);

            mYear = Convert.ToString(Convert.ToInt32(MainServer.OperationDate.ToString("yyyy")) - 1911);
            mMonth = Convert.ToString(MainServer.OperationDate.ToString("MM"));
            mDay = (MethodeName < LrtTxnClientSwitch.Month_AllRideList) ? Convert.ToString(MainServer.OperationDate.ToString("dd")) : "";
        #endregion

        ExportConfigSetting_New:
            #region 修版設定
            MethodeName_zhTW = MethodeName.ToDescription();
            MethodeName_String = MethodeName.ToString();

            reportFileTemp = MainServer.globalCommond.Stock_ReportFile.Find(x => x.ReportNameEng == MethodeName_String);
            S_R_F = MainServer.globalCommond.Stock_ReportFile.FindAll(x => x.ReportNameEng == MethodeName_String).ToList();
            S_R_F_Index = MainServer.globalCommond.Stock_ReportFile.FindIndex(x => x.ReportNameEng == MethodeName_String);
            MyCommond.WriteLog(ThisReceive, $"找到 {((S_R_F_Index > -1) ? $"{S_R_F.Count}" : "0"),4} 筆資料");
            if (S_R_F_Index == -1) { return false; }

            ReportDateStartRow /*----*/ = Convert.ToInt16((/*----------*/ (reportFileTemp.ReportDateRow == "-") /*-----*/ ? "0" : reportFileTemp.ReportDateRow)); /*-----*/ MyCommond.WriteLog(ThisReceive, "產製日期起始行：" + ReportDateStartRow);
            ReportDateStartColumn /*-*/ = MyCommond.ExcelColumnNum((/*-*/ (reportFileTemp.ReportDateColumn == "-") /*--*/ ? "0" : reportFileTemp.ReportDateColumn)); /*--*/ MyCommond.WriteLog(ThisReceive, "產製日期起始欄：" + ReportDateStartColumn);
            StartRow /*--------------*/ = Convert.ToInt16((/*----------*/ (reportFileTemp.StartRow == "-") /*----------*/ ? "0" : reportFileTemp.StartRow)); /*----------*/ MyCommond.WriteLog(ThisReceive, "輸出資料起始行：" + StartRow);
            StartColumn /*-----------*/ = MyCommond.ExcelColumnNum((/*-*/ (reportFileTemp.StartColumn == "-") /*-------*/ ? "0" : reportFileTemp.StartColumn)); /*-------*/ MyCommond.WriteLog(ThisReceive, "輸出資料起始欄：" + StartColumn);
            EndRow /*----------------*/ = Convert.ToInt16((/*----------*/ (reportFileTemp.EndRow == "-") /*------------*/ ? "0" : reportFileTemp.EndRow)); /*------------*/ MyCommond.WriteLog(ThisReceive, "輸出資料結束行：" + EndRow);
            EndColumn /*-------------*/ = MyCommond.ExcelColumnNum((/*-*/ (reportFileTemp.EndColumn == "-") /*---------*/ ? "0" : reportFileTemp.EndColumn)); /*---------*/ MyCommond.WriteLog(ThisReceive, "輸出資料結束欄：" + EndColumn);
            thisBreakOffRow /*-------*/ = Convert.ToInt16((/*----------*/ (reportFileTemp.BreakOffRow == "-") /*-------*/ ? "0" : reportFileTemp.BreakOffRow)); /*-------*/ MyCommond.WriteLog(ThisReceive, "資料中斷行    ：" + thisBreakOffRow);
            thisBreakOffRowNum /*----*/ = Convert.ToInt16((/*----------*/ (reportFileTemp.BreakOffRowNum == "-") /*----*/ ? "0" : reportFileTemp.BreakOffRowNum)); /*----*/ MyCommond.WriteLog(ThisReceive, "資料中斷行數  ：" + thisBreakOffRowNum);
            thisBreakOffColumn /*----*/ = MyCommond.ExcelColumnNum((/*-*/ (reportFileTemp.BreakOffColumn == "-") /*----*/ ? "0" : reportFileTemp.BreakOffColumn)); /*----*/ MyCommond.WriteLog(ThisReceive, "資料中斷欄    ：" + thisBreakOffColumn);
            thisBreakOffColumnNum /*-*/ = Convert.ToInt16((/*----------*/ (reportFileTemp.BreakOffColumnNum == "-") /*-*/ ? "0" : reportFileTemp.BreakOffColumnNum)); /*-*/ MyCommond.WriteLog(ThisReceive, "資料中斷欄數  ：" + thisBreakOffColumnNum);
            MobilePayColumn /*-------*/ = MyCommond.ExcelColumnNum((/*-*/ (reportFileTemp.MobilePayColumn == "-") /*---*/ ? "0" : reportFileTemp.MobilePayColumn)); /*---*/ MyCommond.WriteLog(ThisReceive, "電子支付起始欄：" + MobilePayColumn);
            OwnTiketColumn /*--------*/ = MyCommond.ExcelColumnNum((/*-*/ (reportFileTemp.OwnTiketColumn == "-") /*----*/ ? "0" : reportFileTemp.OwnTiketColumn)); /*----*/ MyCommond.WriteLog(ThisReceive, "自有票起始欄  ：" + OwnTiketColumn);

            mYear = Convert.ToString(Convert.ToInt32(MainServer.OperationDate.ToString("yyyy")) - 1911);
            mMonth = Convert.ToString(MainServer.OperationDate.ToString("MM"));
            mDay = (MethodeName < LrtTxnClientSwitch.Month_AllRideList) ? Convert.ToString(MainServer.OperationDate.ToString("dd")) : "";
            #endregion

            #endregion

            if (Customized) { goto CustomizedSub; } else { goto NormalSub; }

        CustomizedSub:
            MyCommond.WriteLog(ThisReceive, $"客製化運量、營收輸出 - 開始");

            #region 運量、營收

            while (MainServer.VA_Lock) { Thread.Sleep(1000); }

            int Reserve_Year = 1;

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

            #region Offset

            YearAgo = MainServer.OperationDate.AddDays(-0).Year - MainServer.globalCommond.Operation_First_Day.Year;
            MyCommond.WriteLog(ThisReceive, $"變數YearAgo狀態          :{Convert.ToString(YearAgo)}");
            LastMonthDayCount = Convert.ToInt32(DateTime.DaysInMonth(MainServer.OperationDate.AddDays(-1).Year, MainServer.OperationDate.AddDays(-1).Month));  //計算跨月前的月份天數
            MyCommond.WriteLog(ThisReceive, $"變數LastMonthDayCount狀態:{Convert.ToString(LastMonthDayCount)}");
            ThisDay = Convert.ToInt16(MainServer.OperationDate.AddDays(-0).ToString("dd"));
            MyCommond.WriteLog(ThisReceive, $"變數ThisDay狀態          :{Convert.ToString(ThisDay)}");

            IsFirstYear = MainServer.OperationDate.AddDays(-0).Year == MainServer.globalCommond.Operation_First_Day.Year;
            MyCommond.WriteLog(ThisReceive, $"變數IsFirstYear狀態      :{Convert.ToString(IsFirstYear)}");
            IsNextYear = MainServer.OperationDate.AddDays(-1).Year != MainServer.OperationDate.AddDays(-0).Year;
            MyCommond.WriteLog(ThisReceive, $"變數IsNextYear狀態       :{Convert.ToString(IsNextYear)}");
            IsNextMonth = MainServer.OperationDate.AddDays(-1).Month != MainServer.OperationDate.AddDays(-0).Month;
            MyCommond.WriteLog(ThisReceive, $"變數IsNextMonth狀態      :{Convert.ToString(IsNextMonth)}");

            MonthAgo = MainServer.OperationDate.AddDays(-1).Month;
            MyCommond.WriteLog(ThisReceive, $"變數MonthAgo起始狀態     :{Convert.ToString(MonthAgo)}");
            if (IsFirstYear)
            {
                MonthAgo = MonthAgo - MainServer.globalCommond.Operation_First_Day.Month + 1;
                MyCommond.WriteLog(ThisReceive, $"變數MonthAgo狀態         :{Convert.ToString(MonthAgo)}");
            }
            MonthAgo--;

            if (YearAgo != 0) { YearAgo = 1; }

            #endregion

            #region BookSet
            string FileName_DayRideAmt = @"{0}_營運資料統計表(總運量、總營收)_{1}.xlsx";         //1120101_營運資料統計表(總運量、總營收).xlsx

            string BookFileName, BookYear, BookMonth, BookDay;
            string NewBookFileName, NewBookYear, NewBookMonth, NewBookDay;

            //設定開啟檔名
            BookYear = Convert.ToString(Convert.ToInt16(MainServer.OperationDate.AddDays(-1).ToString("yyyy")) - 1911);
            BookMonth = Convert.ToString(MainServer.OperationDate.AddDays(-1).ToString("MM"));
            BookDay = Convert.ToString(MainServer.OperationDate.AddDays(-1).ToString("dd"));
            BookFileName = String.Format(FileName_DayRideAmt, BookYear + BookMonth + BookDay, MainServer.globalCommond.OperationCode);

            NewBookYear = Convert.ToString(Convert.ToInt16(MainServer.OperationDate.AddDays(0).ToString("yyyy")) - 1911);
            NewBookMonth = Convert.ToString(MainServer.OperationDate.AddDays(0).ToString("MM"));
            NewBookDay = Convert.ToString(MainServer.OperationDate.AddDays(0).ToString("dd"));
            NewBookFileName = String.Format(FileName_DayRideAmt, NewBookYear + NewBookMonth + NewBookDay, MainServer.globalCommond.OperationCode);

            MyCommond.WriteLog(ThisReceive, $"變數BookFileName狀態     :{Convert.ToString(BookFileName)}");
            MyCommond.WriteLog(ThisReceive, $"變數NewBookFileName狀態  :{Convert.ToString(NewBookFileName)}");
            #endregion

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
                    int tgC = StartColumn;
                    string msgStr;

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
                            // =SUM(E12:J12,L12:Q12,S12:W12)
                            var S_R_F_Line = MainServer.globalCommond.Stock_ReportFile.FindAll(x => x.ReportNameEng == MethodeName_String);

                            string ForString = "";
                            foreach (var item in S_R_F_Line)
                            {
                                ForString += $"{((ForString != "") ? "," : "")}{item.StartColumn}{tgR}:{item.EndColumn}:{tgR}";
                            }
                            msgStr = $"=SUM({ForString})";

                            ws2.Cells[tgR, tgC] = msgStr;
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
                    MyCommond.WriteLog(ThisReceive, $"在 {MyCommond.ExcelColumnStr(StartColumn)}:{Convert.ToString(tgR + 1)} 寫入總和 {Convert.ToString(MainServer.OperationDate.AddDays(-0).ToString("yyyy/MM/dd"))}");
                    ws2.Cells[tgR + 1, StartColumn] = MainServer.OperationDate.AddDays(-0).ToString("yyyy/MM/dd");
                    MyCommond.WriteLog(ThisReceive, $"月合併結束");
                }
                #endregion

                #region 年合併
                if (IsNextYear)
                {
                    string msgStr;
                    // if (YearAgo > Reserve_Year) { RowOffset = Reserve_Year; } else { RowOffset = YearAgo; }
                    MyCommond.WriteLog(ThisReceive, $"年合併開始");
                    int tgR = StartRow + RowOffset - 1;
                    int tgC = StartColumn;
                    msgStr = "將{0}:{1}寫到{0}:{2}";

                    //填寫或搬移年總運量(保留近3年)
                    while (tgC <= EndColumn)
                    {
                        if (tgC == StartColumn)
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
                            // =SUM(E12:J12,L12:Q12,S12:W12)
                            var S_R_F_Line = MainServer.globalCommond.Stock_ReportFile.FindAll(x => x.ReportNameEng == MethodeName_String);

                            string ForString = "";
                            foreach (var item in S_R_F_Line)
                            {
                                ForString += $"{((ForString != "") ? "," : "")}{item.StartColumn}{tgR}:{item.EndColumn}:{tgR}";
                            }
                            msgStr = $"=SUM({ForString})";

                            ws2.Cells[tgR, tgC] = msgStr;
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

                    MyCommond.WriteLog(ThisReceive, $"開始移除每月運量");
                    //移除每月總運量
                    Int32 WrithTime = 0;
                    strRange = (tgR + 1) + ":" + ((tgR + 1) + MonthAgo);
                    ws2.Rows[strRange].delete();
                }
                #endregion

                #region 合併資料
                MyCommond.WriteLog(ThisReceive, $"合併空間中的資料");
                var Line1 = MainServer.globalCommond.Stock_ReportFile.Find(x =>
                    x.ReportNameEng == MethodeName_String &&
                    x.OperationLine == MainServer.globalCommond.OperationCode);

                var Merge = Data_Mrt;

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

                IsFirstYear = MainServer.OperationDate.AddDays(-0).Year == MainServer.globalCommond.Operation_First_Day.Year;
                MyCommond.WriteLog(ThisReceive, $"變更IsFirstYear狀態      :{Convert.ToString(IsFirstYear)}");

                MonthAgo = Convert.ToDateTime(MainServer.OperationDate.ToString("yyyy/MM/01")).AddMonths(-1).Month;
                if (MonthAgo == 12) { MonthAgo = 0; }

                if (IsFirstYear) MonthAgo = MonthAgo - MainServer.globalCommond.Operation_First_Day.Month;
                MyCommond.WriteLog(ThisReceive, $"變數MonthAgo狀態:{Convert.ToString(MonthAgo)}");

                //if (IsFirstYear) WriteRow = StartRow + MonthAgo + ThisDay;
                //else if (YearAgo > Reserve_Year) WriteRow = StartRow + Reserve_Year + (MonthAgo - 1) + ThisDay;
                //else WriteRow = StartRow + YearAgo + (MonthAgo - 1) + ThisDay;

                WriteRow = StartRow + MonthAgo + ThisDay - 1;
                if (!IsFirstYear) WriteRow += YearAgo;

                MyCommond.WriteLog(ThisReceive, $"寫入起始行:{WriteRow}");

                string msgStr_Y = "第 {0} 行第 {1} 欄寫入第 {2} 筆資料：{3}";
                WriteColumn = MyCommond.ExcelColumnNum(Line1.StartColumn);
                MyCommond.WriteLog(ThisReceive, $"寫入起始欄:{WriteColumn}");
                for (int i = 0; i < Merge.GetLength(0); i++)
                {
                    for (int ColumnOffset2 = 0; ColumnOffset2 < Merge.GetLength(1); ColumnOffset2++)
                    {
                        int s1 = WriteRow;
                        int s2 = WriteColumn + ColumnOffset2;
                        int s3 = (i + 1) * (ColumnOffset2 + 1);
                        int s4 = Merge[i, ColumnOffset2];
                        MyCommond.WriteLog(ThisReceive, $"{string.Format(msgStr_Y, s1, s2, s3, s4)}");
                        ws2.Cells[s1, s2] = s4;
                    }
                }
                MyCommond.WriteLog(ThisReceive, $"");

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

            #endregion

            MyCommond.WriteLog(ThisReceive, $"客製化運量、營收輸出 - 結束");
            return ExportCheck;

        NormalSub:
            MyCommond.WriteLog(ThisReceive, $"其他通用輸出 - 開始");

            #region 普通

            bool ckeck = MyCommond.CheckFile(exportFullPath);
            XLWorkbook wb = new XLWorkbook();
            if (ckeck)
            {
                MyCommond.WriteLog(ThisReceive, "以既有的檔案複寫");
                wb = new XLWorkbook(exportFullPath);
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

            try
            {
                var ws = wb.Worksheets.First();
                int RowBreak, ColumnBreak;

                MyCommond.WriteLog(ThisReceive, "寫入產製日期");
                ws.Cell(ReportDateStartRow, ReportDateStartColumn).Value = $"{MainServer.OperationDate}";
                ws.Cell(ReportDateStartRow + 1, ReportDateStartColumn).Value = DateTime.Now.ToString("yyyy/MM/dd");

                #region 一般資料寫入

                try
                {
                    MyCommond.WriteLog(ThisReceive, $"Data_Afc -> Row:{Data_Mrt.GetLength(0),4}, Column:{Data_Mrt.GetLength(1),4}");
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"Data_Afc -> Error\n{ex.Message}");
                }
                finally
                {

                }

                if (Data_Mrt != null && Data_Mrt.GetLength(0) > 0 && Data_Mrt.GetLength(1) > 0)
                {
                    if (thisBreakOffRow != 0)
                    {
                        RowBreak = thisBreakOffRow - StartRow + 1;
                        EndRow = EndRow - StartRow - thisBreakOffRowNum;
                    }
                    else
                    {
                        RowBreak = Data_Mrt.GetLength(0);
                    }
                    if (thisBreakOffColumn != 0)
                    {
                        ColumnBreak = thisBreakOffColumn - StartColumn + 1;
                        EndColumn = EndColumn - StartColumn - thisBreakOffColumnNum + 1;
                    }
                    else
                    {
                        ColumnBreak = Data_Mrt.GetLength(1);
                    }

                    MyCommond.WriteLog(ThisReceive, $"---RowBreak:{RowBreak,4}, EndRow:{EndRow,4}, ColumnBreak:{ColumnBreak,4}, EndColumn:{EndColumn,4}---");

                    MyCommond.WriteLog(ThisReceive, "開始資料寫入");
                    for (int RowOffset_1 = 0; RowOffset_1 < RowBreak; RowOffset_1++)
                    {
                        try
                        {
                            for (int ColumnOffset = 0; ColumnOffset < ColumnBreak; ColumnOffset++)
                            {
                                MyCommond.WriteLog(ThisReceive, $"" +
                                    $"sheet.Cell({(StartRow + RowOffset_1),4}, {(StartColumn + ColumnOffset),4}).Value = " +
                                    $"{Data_Mrt[RowOffset_1, ColumnOffset]}");
                                ws.Cell(StartRow + RowOffset_1, StartColumn + ColumnOffset).Value = Data_Mrt[RowOffset_1, ColumnOffset];
                            }
                            if (thisBreakOffColumn != 0)
                            {
                                for (int ColumnOffset = ColumnBreak; ColumnOffset < EndColumn; ColumnOffset++)
                                {
                                    MyCommond.WriteLog(ThisReceive, $"" +
                                        $"sheet.Cell({(StartRow + RowOffset_1),4}, {(StartColumn + thisBreakOffColumnNum + ColumnOffset),4}).Value = " +
                                        $"{Data_Mrt[RowOffset_1, ColumnOffset]}");
                                    ws.Cell(StartRow + RowOffset_1, StartColumn + thisBreakOffColumnNum + ColumnOffset).Value = Data_Mrt[RowOffset_1, ColumnOffset];
                                }
                            }//斷點欄後寫數值不對
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
                                        $"{Data_Mrt[RowOffset_2, ColumnOffset]}");
                                    ws.Cell(StartRow + thisBreakOffRowNum + RowOffset_2, StartColumn + ColumnOffset).Value = Data_Mrt[RowOffset_2, ColumnOffset];
                                }
                                if (thisBreakOffColumn != 0)
                                {
                                    for (int ColumnOffset = ColumnBreak; ColumnOffset < EndColumn; ColumnOffset++)
                                    {
                                        MyCommond.WriteLog(ThisReceive, $"" +
                                            $"sheet.Cell({(StartRow + thisBreakOffRowNum + RowOffset_2),4}, {(StartColumn + thisBreakOffColumnNum + ColumnOffset),4}).Value = " +
                                            $"{Data_Mrt[RowOffset_2, ColumnOffset]}");
                                        ws.Cell(StartRow + thisBreakOffRowNum + RowOffset_2, StartColumn + thisBreakOffColumnNum + ColumnOffset).Value = Data_Mrt[RowOffset_2, ColumnOffset];
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                MyCommond.WriteLog(ThisReceive, $"ERROR! \n{MyCommond.ExportErrorMessageToLog(ex, MethodeName.ToDescription()).Message}");
                            }
                        }
                    }

                    if (Data_Other != null)
                    {
                        int OtherInputOffset = StartRow + Data_Mrt.GetLength(0) + 2;
                        for (int columnOffset = 0; columnOffset < Data_Other.Length; columnOffset++)
                        {
                            MyCommond.WriteLog(ThisReceive, $"" +
                                 $"sheet.Cell({OtherInputOffset,4}, {(StartColumn + columnOffset),4}).Value = " +
                                 $"{Data_Other[columnOffset]}");
                            ws.Cell(OtherInputOffset, StartColumn + columnOffset).Value = Data_Other[columnOffset];
                        }
                    }
                }

                

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
                            int ExcelCellValue = ((ws.Cell(RowSet, ColumnSet).Value == "") ? 0 : Convert.ToInt32(ws.Cell(RowSet, ColumnSet).Value));
                            MyCommond.WriteLog(ThisReceive, $"" +
                                $"sheet.Cell({RowSet}, {ColumnSet}).Value = " +
                                $"{ExcelCellValue} + {Data_MobilePay[RowOffset_3, ColumnOffset_3]} = " +
                                $"{ExcelCellValue + Data_MobilePay[RowOffset_3, ColumnOffset_3]}");
                            ws.Cell(RowSet, ColumnSet).Value = Convert.ToInt64(ExcelCellValue) + Data_MobilePay[RowOffset_3, ColumnOffset_3];
                        }
                    }
                }

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

                #endregion

                MyCommond.WriteLog(ThisReceive, $"寫入結束，將檔案存成{ec.ToString()}");
                ExportCheck = true;
                
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR! \n{MyCommond.ExportErrorMessageToLog(ex, MethodeName.ToDescription()).Message}");
            }
            finally
            {
                wb.SaveAs(exportFullPath);
                wb = null;
                MyCommond.WriteLog(ThisReceive, $"結束Excel輸出");
            }

            #endregion

            MyCommond.WriteLog(ThisReceive, $"其他通用輸出 - 結束");

            return ExportCheck;
        }

        public bool ExportExcle(ExportConfig ec, LrtTxnClientSwitch MethodeName, double[,] Data_Mrt, double[,] Data_MobilePay, double[,] Data_OwnTicket, double[] Data_Other = null)
        {
            lock (MainServer)
            {

            string exportFullPath = $"{ec.Path1}{ec.ToString()}";
            string template = MyCommond.Path_Program + MyCommond.Path_Template + MethodeName.ToString() + ".xlsx";

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

            mYear = Convert.ToString(Convert.ToInt32(MainServer.OperationDate.ToString("yyyy")) - 1911);
            mMonth = Convert.ToString(MainServer.OperationDate.ToString("MM"));
            mDay = (MethodeName < LrtTxnClientSwitch.Month_AllRideList) ? Convert.ToString(MainServer.OperationDate.ToString("dd")) : "";

            #endregion 讀取輸出位置資料

            if (Customized) { goto CustomizedSub; }
            else { goto NormalSub; }

        CustomizedSub:
            MyCommond.WriteLog(ThisReceive, $"客製化運量、營收輸出 - 開始");

            #region 運量、營收
            
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

            #region Offset

            YearAgo = MainServer.OperationDate.AddDays(-0).Year - MainServer.globalCommond.Operation_First_Day.Year;
            MyCommond.WriteLog(ThisReceive, $"變數YearAgo狀態          :{Convert.ToString(YearAgo)}");
            LastMonthDayCount = Convert.ToInt32(DateTime.DaysInMonth(MainServer.OperationDate.AddDays(-1).Year, MainServer.OperationDate.AddDays(-1).Month));  //計算跨月前的月份天數
            MyCommond.WriteLog(ThisReceive, $"變數LastMonthDayCount狀態:{Convert.ToString(LastMonthDayCount)}");
            ThisDay = Convert.ToInt16(MainServer.OperationDate.AddDays(-0).ToString("dd"));
            MyCommond.WriteLog(ThisReceive, $"變數ThisDay狀態          :{Convert.ToString(ThisDay)}");

            IsFirstYear = MainServer.OperationDate.AddDays(-1).Year == MainServer.globalCommond.Operation_First_Day.Year;
            MyCommond.WriteLog(ThisReceive, $"變數IsFirstYear狀態      :{Convert.ToString(IsFirstYear)}");
            IsNextYear = MainServer.OperationDate.AddDays(-1).Year != MainServer.OperationDate.AddDays(-0).Year;
            MyCommond.WriteLog(ThisReceive, $"變數IsNextYear狀態       :{Convert.ToString(IsNextYear)}");
            IsNextMonth = MainServer.OperationDate.AddDays(-1).Month != MainServer.OperationDate.AddDays(-0).Month;
            MyCommond.WriteLog(ThisReceive, $"變數IsNextMonth狀態      :{Convert.ToString(IsNextMonth)}");

            MonthAgo = MainServer.OperationDate.AddDays(-1).Month;
            MyCommond.WriteLog(ThisReceive, $"變數MonthAgo起始狀態     :{Convert.ToString(MonthAgo)}");
            if (IsFirstYear)
            {
                MonthAgo = MonthAgo - MainServer.globalCommond.Operation_First_Day.Month + 1;
                MyCommond.WriteLog(ThisReceive, $"變數MonthAgo狀態         :{Convert.ToString(MonthAgo)}");
            }
            MonthAgo--;

            if (YearAgo != 0) { YearAgo = 1; }

            #endregion

            string FileName_DayRideAmt = @"{0}_營運資料統計表(總運量、總營收)_{1}.xlsx";         //1120101_營運資料統計表(總運量、總營收).xlsx

            string BookFileName, BookYear, BookMonth, BookDay;
            string NewBookFileName, NewBookYear, NewBookMonth, NewBookDay;

            var S_R_F_Line = MainServer.globalCommond.Stock_ReportFile.FindAll(x =>
                x.ReportNameEng == MethodeName_String
                && x.OperationLine == MainServer.globalCommond.OperationCode
                );

            //設定開啟檔名
            BookYear = Convert.ToString(Convert.ToInt16(MainServer.OperationDate.AddDays(-1).ToString("yyyy")) - 1911);
            BookMonth = Convert.ToString(MainServer.OperationDate.AddDays(-1).ToString("MM"));
            BookDay = Convert.ToString(MainServer.OperationDate.AddDays(-1).ToString("dd"));
            BookFileName = String.Format(FileName_DayRideAmt, BookYear + BookMonth + BookDay, MainServer.globalCommond.OperationCode_String);

            NewBookYear = Convert.ToString(Convert.ToInt16(MainServer.OperationDate.AddDays(0).ToString("yyyy")) - 1911);
            NewBookMonth = Convert.ToString(MainServer.OperationDate.AddDays(0).ToString("MM"));
            NewBookDay = Convert.ToString(MainServer.OperationDate.AddDays(0).ToString("dd"));
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
                    else if (YearAgo > Reserve_Year) { RowOffset = Reserve_Year + 1 + MonthAgo; }
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

                            ws2.Cells[tgR, tgC] = string.Format("=SUM({0})", ForString);
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
                    MyCommond.WriteLog(ThisReceive, $"在 {S_R_F_Line[0].ReportDateColumn}:{Convert.ToString(tgR + 1)} 寫入總和 {Convert.ToString(MainServer.OperationDate.AddDays(-0).ToString("yyyy/MM/dd"))}");
                    ws2.Cells[tgR + 1, MyCommond.ExcelColumnNum(S_R_F_Line[0].ReportDateColumn)] = MainServer.OperationDate.AddDays(-0).ToString("yyyy/MM/dd");
                    MyCommond.WriteLog(ThisReceive, $"月合併結束");
                }
                #endregion

                #region 年合併
                if (IsNextYear)
                {
                    string msgStr;
                    // if (YearAgo > Reserve_Year) { RowOffset = Reserve_Year; } else { RowOffset = YearAgo; }
                    MyCommond.WriteLog(ThisReceive, $"年合併開始");
                    int tgR = StartRow;
                    int tgC = ReportDateStartColumn;
                    msgStr = "將{0}:{1}寫到{0}:{2}";

                    //填寫或搬移年總運量(保留近3年)
                    while (tgC <= EndColumn)
                    {
                        if (tgC == ReportDateStartColumn)
                        {
                            if (YearAgo < Reserve_Year)
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
                                //ws2.Cells[tgR - 2, ReportDateStartColumn] = ws2.Cells[tgR - 1, ReportDateStartColumn];
                                //ws2.Cells[tgR - 1, ReportDateStartColumn] = ws2.Cells[tgR + 0, ReportDateStartColumn];
                                ws2.Cells[tgR + 0, ReportDateStartColumn] = ws2.Cells[tgR + 1, ReportDateStartColumn];
                            }
                            //取得日期
                            StrDate = Convert.ToDateTime(ws2.Cells[tgR + 1, tgC].text);
                            //將日期重新編排輸入
                            ws2.Cells[tgR, tgC].NumberFormatLocal = "yyyy年";
                            ws2.Cells[tgR, tgC] = StrDate.ToString("yyyy/01/01");
                            //msgStr = "取得 " + MyCommond.ExcelColumnStr(tgC) + ":" + Convert.ToString(tgR + 1) + " 資料 " + StrDate + " 寫入 " + MyCommond.ExcelColumnStr(tgC) + Convert.ToString(tgR) + "並將設改為YYYY.MM";
                            msgStr = string.Format("取得{0}{1}資料{2}寫入{3}{4}，並將資料型態設為yyyy.MM", MyCommond.ExcelColumnStr(tgC), Convert.ToString(tgR + 1), StrDate, MyCommond.ExcelColumnStr(tgC), Convert.ToString(tgR));
                            MyCommond.WriteLog(ThisReceive, $"{msgStr}");
                        }   //寫入該月份
                        else if (tgC == ReportDateStartColumn + 1)
                        {
                            MyCommond.WriteLog(ThisReceive, $"此欄不做處理");
                        }   //星期
                        else if (tgC == ReportDateStartColumn + 2)
                        {
                            // =SUM(E12:J12,L12:Q12,S12:W12)

                            string ForString = "";
                            var Totle_Formula = MainServer.globalCommond.Stock_ReportFile.FindAll(x =>
                                x.ReportNameEng == MethodeName_String
                                );
                                foreach (var item in Totle_Formula)
                            {
                                ForString += $"{((ForString != "") ? "," : "")}{item.StartColumn}{tgR}:{item.EndColumn}{tgR}";
                            }
                            msgStr = $"=SUM({ForString})";

                            ws2.Cells[tgR, tgC] = msgStr;
                            MyCommond.WriteLog(ThisReceive, $"{msgStr}");
                        }   //總計
                        else if (tgC >= StartColumn && tgC <= EndColumn)
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

                    MyCommond.WriteLog(ThisReceive, $"開始移除每月運量");
                    //移除每月總運量
                    Int32 WrithTime = 0;
                    strRange = (tgR + 1) + ":" + ((tgR + 1) + MonthAgo);
                    ws2.Rows[strRange].delete();
                }
                #endregion

                #region 合併資料
                MyCommond.WriteLog(ThisReceive, $"合併空間中的資料");
                var Line1 = MainServer.globalCommond.Stock_ReportFile.Find(x =>
                    x.ReportNameEng == MethodeName_String
                    && x.OperationLine == MainServer.globalCommond.OperationCode
                    );

                var Merge = Data_Mrt;

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

                IsFirstYear = MainServer.OperationDate.AddDays(-0).Year == MainServer.globalCommond.Operation_First_Day.Year;
                MyCommond.WriteLog(ThisReceive, $"變更IsFirstYear狀態      :{Convert.ToString(IsFirstYear)}");

                MonthAgo = Convert.ToDateTime(MainServer.OperationDate.ToString("yyyy/MM/01")).AddMonths(-1).Month;
                if (MonthAgo == 12) { MonthAgo = 0; }

                if (IsFirstYear) MonthAgo = MonthAgo - MainServer.globalCommond.Operation_First_Day.Month;
                MyCommond.WriteLog(ThisReceive, $"變數MonthAgo狀態:{Convert.ToString(MonthAgo)}");

                //if (IsFirstYear) WriteRow = StartRow + MonthAgo + ThisDay;
                //else if (YearAgo > Reserve_Year) WriteRow = StartRow + Reserve_Year + (MonthAgo - 1) + ThisDay;
                //else WriteRow = StartRow + YearAgo + (MonthAgo - 1) + ThisDay;

                WriteRow = StartRow + MonthAgo + ThisDay - 1;
                if (!IsFirstYear) WriteRow += YearAgo;

                MyCommond.WriteLog(ThisReceive, $"寫入起始行:{WriteRow}");

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
                        double s4 = Merge[i, ColumnOffset2];
                        MyCommond.WriteLog(ThisReceive, $"{string.Format(msgStr_Y, s1, s2, sss++, s4)}");
                        ws2.Cells[s1, s2] = s4;
                    }
                }
                MyCommond.WriteLog(ThisReceive, $"");

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

            #endregion 運量、營收

            MyCommond.WriteLog(ThisReceive, $"客製化運量、營收輸出 - 結束");
            return ExportCheck;

        NormalSub:
            MyCommond.WriteLog(ThisReceive, $"其他通用輸出 - 開始");

            #region 普通

            bool ckeck = MyCommond.CheckFile(exportFullPath);
            XLWorkbook wb = new XLWorkbook();
            if (ckeck)
            {
                MyCommond.WriteLog(ThisReceive, "以既有的檔案複寫");
                wb = new XLWorkbook(exportFullPath);
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

            try
            {
                var ws = wb.Worksheets.First();
                int RowBreak, ColumnBreak;

                MyCommond.WriteLog(ThisReceive, "寫入產製日期");
                ws.Cell(ReportDateStartRow, ReportDateStartColumn).Value = $"{MainServer.OperationDate}";
                ws.Cell(ReportDateStartRow + 1, ReportDateStartColumn).Value = DateTime.Now.ToString("yyyy/MM/dd");

                #region 一般資料寫入

                try
                {
                    MyCommond.WriteLog(ThisReceive, $"Data_Afc -> Row:{Data_Mrt.GetLength(0),4}, Column:{Data_Mrt.GetLength(1),4}");
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"Data_Afc -> Error\n{ex.Message}");
                }
                finally
                {

                }

                if (Data_Mrt != null && Data_Mrt.GetLength(0) > 0 && Data_Mrt.GetLength(1) > 0)
                {
                    if (thisBreakOffRow != 0)
                    {
                        RowBreak = thisBreakOffRow - StartRow + 1;
                        EndRow = EndRow - StartRow - thisBreakOffRowNum;
                    }
                    else
                    {
                        RowBreak = Data_Mrt.GetLength(0);
                    }
                    if (thisBreakOffColumn != 0)
                    {
                        ColumnBreak = thisBreakOffColumn - StartColumn + 1;
                        EndColumn = EndColumn - StartColumn - thisBreakOffColumnNum + 1;
                    }
                    else
                    {
                        ColumnBreak = Data_Mrt.GetLength(1);
                    }

                    MyCommond.WriteLog(ThisReceive, $"---RowBreak:{RowBreak,4}, EndRow:{EndRow,4}, ColumnBreak:{ColumnBreak,4}, EndColumn:{EndColumn,4}---");

                    MyCommond.WriteLog(ThisReceive, "開始資料寫入");
                    for (int RowOffset_1 = 0; RowOffset_1 < RowBreak; RowOffset_1++)
                    {
                        try
                        {
                            for (int ColumnOffset = 0; ColumnOffset < ColumnBreak; ColumnOffset++)
                            {
                                MyCommond.WriteLog(ThisReceive, $"" +
                                    $"sheet.Cell({(StartRow + RowOffset_1),4}, {(StartColumn + ColumnOffset),4}).Value = " +
                                    $"{Data_Mrt[RowOffset_1, ColumnOffset]}");
                                ws.Cell(StartRow + RowOffset_1, StartColumn + ColumnOffset).Value = Data_Mrt[RowOffset_1, ColumnOffset];
                            }
                            if (thisBreakOffColumn != 0)
                            {
                                for (int ColumnOffset = ColumnBreak; ColumnOffset < EndColumn; ColumnOffset++)
                                {
                                    MyCommond.WriteLog(ThisReceive, $"" +
                                        $"sheet.Cell({(StartRow + RowOffset_1),4}, {(StartColumn + thisBreakOffColumnNum + ColumnOffset),4}).Value = " +
                                        $"{Data_Mrt[RowOffset_1, ColumnOffset]}");
                                    ws.Cell(StartRow + RowOffset_1, StartColumn + thisBreakOffColumnNum + ColumnOffset).Value = Data_Mrt[RowOffset_1, ColumnOffset];
                                }
                            }//斷點欄後寫數值不對
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
                                        $"{Data_Mrt[RowOffset_2, ColumnOffset]}");
                                    ws.Cell(StartRow + thisBreakOffRowNum + RowOffset_2, StartColumn + ColumnOffset).Value = Data_Mrt[RowOffset_2, ColumnOffset];
                                }
                                if (thisBreakOffColumn != 0)
                                {
                                    for (int ColumnOffset = ColumnBreak; ColumnOffset < EndColumn; ColumnOffset++)
                                    {
                                        MyCommond.WriteLog(ThisReceive, $"" +
                                            $"sheet.Cell({(StartRow + thisBreakOffRowNum + RowOffset_2),4}, {(StartColumn + thisBreakOffColumnNum + ColumnOffset),4}).Value = " +
                                            $"{Data_Mrt[RowOffset_2, ColumnOffset]}");
                                        ws.Cell(StartRow + thisBreakOffRowNum + RowOffset_2, StartColumn + thisBreakOffColumnNum + ColumnOffset).Value = Data_Mrt[RowOffset_2, ColumnOffset];
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                MyCommond.WriteLog(ThisReceive, $"ERROR! \n{MyCommond.ExportErrorMessageToLog(ex, MethodeName.ToDescription()).Message}");
                            }
                        }
                    }

                    if (Data_Other != null)
                    {
                        int OtherInputOffset = StartRow + Data_Mrt.GetLength(0) + 2;
                        for (int columnOffset = 0; columnOffset < Data_Other.Length; columnOffset++)
                        {
                            MyCommond.WriteLog(ThisReceive, $"" +
                                 $"sheet.Cell({OtherInputOffset,4}, {(StartColumn + columnOffset),4}).Value = " +
                                 $"{Data_Other[columnOffset]}");
                            ws.Cell(OtherInputOffset, StartColumn + columnOffset).Value = Data_Other[columnOffset];
                        }
                    }
                }



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
                            int ExcelCellValue = ((ws.Cell(RowSet, ColumnSet).Value == "") ? 0 : Convert.ToInt32(ws.Cell(RowSet, ColumnSet).Value));
                            MyCommond.WriteLog(ThisReceive, $"" +
                                $"sheet.Cell({RowSet}, {ColumnSet}).Value = " +
                                $"{ExcelCellValue} + {Data_MobilePay[RowOffset_3, ColumnOffset_3]} = " +
                                $"{ExcelCellValue + Data_MobilePay[RowOffset_3, ColumnOffset_3]}");
                            ws.Cell(RowSet, ColumnSet).Value = Convert.ToInt64(ExcelCellValue) + Data_MobilePay[RowOffset_3, ColumnOffset_3];
                        }
                    }
                }

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

                #endregion

                MyCommond.WriteLog(ThisReceive, $"寫入結束，將檔案存成{ec.ToString()}");
                ExportCheck = true;

            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR! \n{MyCommond.ExportErrorMessageToLog(ex, MethodeName.ToDescription()).Message}");
            }
            finally
            {
                wb.SaveAs(exportFullPath);
                wb = null;
                MyCommond.WriteLog(ThisReceive, $"結束Excel輸出");
            }

            #endregion

            MyCommond.WriteLog(ThisReceive, $"其他通用輸出 - 結束");

            return ExportCheck;
            }
        }

        public bool ExportExcle(ExportConfig ec, MrtTxnClientSwitch MethodeName, double[,] Data_Mrt, double[,] Data_MobilePay, double[,] Data_OwnTicket, double[] Data_Other = null)
        {
            lock (MainServer)
            {

                string exportFullPath = $"{ec.Path1}{ec.ToString()}";
                string template = MyCommond.Path_Program + MyCommond.Path_Template + MethodeName.ToString() + ".xlsx";

                #region 讀取輸出位置資料

                MyCommond.WriteLog(ThisReceive, $"取得輸出資料變數");
                string mYear = "", mMonth = "", mDay = "";

                bool ExportCheck = false;
                

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

                mYear = Convert.ToString(Convert.ToInt32(MainServer.OperationDate.ToString("yyyy")) - 1911);
                mMonth = Convert.ToString(MainServer.OperationDate.ToString("MM"));
                mDay = Convert.ToString(MainServer.OperationDate.ToString("dd"));

                #endregion 讀取輸出位置資料

                MyCommond.WriteLog(ThisReceive, $"其他通用輸出 - 開始");

                #region 普通

                bool ckeck = MyCommond.CheckFile(exportFullPath);
                XLWorkbook wb = new XLWorkbook();
                if (ckeck)
                {
                    MyCommond.WriteLog(ThisReceive, "以既有的檔案複寫");
                    wb = new XLWorkbook(exportFullPath);
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

                try
                {
                    var ws = wb.Worksheets.First();
                    int RowBreak, ColumnBreak;

                    MyCommond.WriteLog(ThisReceive, "寫入產製日期");
                    ws.Cell(ReportDateStartRow, ReportDateStartColumn).Value = $"{MainServer.OperationDate}";
                    ws.Cell(ReportDateStartRow + 1, ReportDateStartColumn).Value = DateTime.Now.ToString("yyyy/MM/dd");

                    #region 一般資料寫入

                    try
                    {
                        MyCommond.WriteLog(ThisReceive, $"Data_Afc -> Row:{Data_Mrt.GetLength(0),4}, Column:{Data_Mrt.GetLength(1),4}");
                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $"Data_Afc -> Error\n{ex.Message}");
                    }
                    finally
                    {

                    }

                    if (Data_Mrt != null && Data_Mrt.GetLength(0) > 0 && Data_Mrt.GetLength(1) > 0)
                    {
                        if (thisBreakOffRow != 0)
                        {
                            RowBreak = thisBreakOffRow - StartRow + 1;
                            EndRow = EndRow - StartRow - thisBreakOffRowNum;
                        }
                        else
                        {
                            RowBreak = Data_Mrt.GetLength(0);
                        }
                        if (thisBreakOffColumn != 0)
                        {
                            ColumnBreak = thisBreakOffColumn - StartColumn + 1;
                            EndColumn = EndColumn - StartColumn - thisBreakOffColumnNum + 1;
                        }
                        else
                        {
                            ColumnBreak = Data_Mrt.GetLength(1);
                        }

                        MyCommond.WriteLog(ThisReceive, $"---RowBreak:{RowBreak,4}, EndRow:{EndRow,4}, ColumnBreak:{ColumnBreak,4}, EndColumn:{EndColumn,4}---");

                        MyCommond.WriteLog(ThisReceive, "開始資料寫入");
                        for (int RowOffset_1 = 0; RowOffset_1 < RowBreak; RowOffset_1++)
                        {
                            try
                            {
                                for (int ColumnOffset = 0; ColumnOffset < ColumnBreak; ColumnOffset++)
                                {
                                    MyCommond.WriteLog(ThisReceive, $"" +
                                        $"sheet.Cell({(StartRow + RowOffset_1),4}, {(StartColumn + ColumnOffset),4}).Value = " +
                                        $"{Data_Mrt[RowOffset_1, ColumnOffset]}");
                                    ws.Cell(StartRow + RowOffset_1, StartColumn + ColumnOffset).Value = Data_Mrt[RowOffset_1, ColumnOffset];
                                }
                                if (thisBreakOffColumn != 0)
                                {
                                    for (int ColumnOffset = ColumnBreak; ColumnOffset < EndColumn; ColumnOffset++)
                                    {
                                        MyCommond.WriteLog(ThisReceive, $"" +
                                            $"sheet.Cell({(StartRow + RowOffset_1),4}, {(StartColumn + thisBreakOffColumnNum + ColumnOffset),4}).Value = " +
                                            $"{Data_Mrt[RowOffset_1, ColumnOffset]}");
                                        ws.Cell(StartRow + RowOffset_1, StartColumn + thisBreakOffColumnNum + ColumnOffset).Value = Data_Mrt[RowOffset_1, ColumnOffset];
                                    }
                                }//斷點欄後寫數值不對
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
                                            $"{Data_Mrt[RowOffset_2, ColumnOffset]}");
                                        ws.Cell(StartRow + thisBreakOffRowNum + RowOffset_2, StartColumn + ColumnOffset).Value = Data_Mrt[RowOffset_2, ColumnOffset];
                                    }
                                    if (thisBreakOffColumn != 0)
                                    {
                                        for (int ColumnOffset = ColumnBreak; ColumnOffset < EndColumn; ColumnOffset++)
                                        {
                                            MyCommond.WriteLog(ThisReceive, $"" +
                                                $"sheet.Cell({(StartRow + thisBreakOffRowNum + RowOffset_2),4}, {(StartColumn + thisBreakOffColumnNum + ColumnOffset),4}).Value = " +
                                                $"{Data_Mrt[RowOffset_2, ColumnOffset]}");
                                            ws.Cell(StartRow + thisBreakOffRowNum + RowOffset_2, StartColumn + thisBreakOffColumnNum + ColumnOffset).Value = Data_Mrt[RowOffset_2, ColumnOffset];
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MyCommond.WriteLog(ThisReceive, $"ERROR! \n{MyCommond.ExportErrorMessageToLog(ex, MethodeName.ToDescription()).Message}");
                                }
                            }
                        }

                        if (Data_Other != null)
                        {
                            int OtherInputOffset = StartRow + Data_Mrt.GetLength(0) + 2;
                            for (int columnOffset = 0; columnOffset < Data_Other.Length; columnOffset++)
                            {
                                MyCommond.WriteLog(ThisReceive, $"" +
                                     $"sheet.Cell({OtherInputOffset,4}, {(StartColumn + columnOffset),4}).Value = " +
                                     $"{Data_Other[columnOffset]}");
                                ws.Cell(OtherInputOffset, StartColumn + columnOffset).Value = Data_Other[columnOffset];
                            }
                        }
                    }



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
                                int ExcelCellValue = ((ws.Cell(RowSet, ColumnSet).Value == "") ? 0 : Convert.ToInt32(ws.Cell(RowSet, ColumnSet).Value));
                                MyCommond.WriteLog(ThisReceive, $"" +
                                    $"sheet.Cell({RowSet}, {ColumnSet}).Value = " +
                                    $"{ExcelCellValue} + {Data_MobilePay[RowOffset_3, ColumnOffset_3]} = " +
                                    $"{ExcelCellValue + Data_MobilePay[RowOffset_3, ColumnOffset_3]}");
                                ws.Cell(RowSet, ColumnSet).Value = Convert.ToInt64(ExcelCellValue) + Data_MobilePay[RowOffset_3, ColumnOffset_3];
                            }
                        }
                    }

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

                    #endregion

                    MyCommond.WriteLog(ThisReceive, $"寫入結束，將檔案存成{ec.ToString()}");
                    ExportCheck = true;

                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"ERROR! \n{MyCommond.ExportErrorMessageToLog(ex, MethodeName.ToDescription()).Message}");
                }
                finally
                {
                    wb.SaveAs(exportFullPath);
                    wb = null;
                    MyCommond.WriteLog(ThisReceive, $"結束Excel輸出");
                }

                #endregion

                MyCommond.WriteLog(ThisReceive, $"其他通用輸出 - 結束");

                return ExportCheck;
            }
        }

        public bool ExportExcel(ExportConfig ec, LrtTxnClientSwitch MethodeName, string[,] Data_Mrt, string[,] Data_MobilePay, string[,] Data_OwnTicket, string[] Data_Other = null)
        {
            double[,] Data_Mrt_2_Int = String2Int(Data_Mrt);
            double[,] Data_MobilePay_2_Int = String2Int(Data_MobilePay);
            double[,] Data_OwnTicket_2_Int = String2Int(Data_OwnTicket);
            double[] Data_Other_2_Int = String2Int(Data_Other);

            ExportExcle(ec, MethodeName, Data_Mrt_2_Int, Data_MobilePay_2_Int, Data_OwnTicket_2_Int, Data_Other_2_Int);

            return false;
        }

        private double[] String2Int(string[] aa)
        {
            double[] re = null;
            if (aa != null)
            {
                if (aa.Length > 0)
                {
                    re = new double[aa.Length];
                    for (int i = 0; i < aa.Length; i++)
                    {
                        re[i] = Convert.ToInt32(aa[i]);
                    }
                }
            }

            return re;
        }

        private double[,] String2Int(string[,] aa)
        {
            double[,] re = null;
            if (aa != null)
            {
                if (aa.GetLength(0) > 0 && aa.GetLength(1) > 0)
                {
                    re = new double[aa.GetLength(0), aa.GetLength(1)];
                    for (int i = 0; i < aa.GetLength(0); i++)
                    {
                        for (int j = 0; j < aa.GetLength(1); j++)
                        {
                            re[i,j] = Convert.ToInt32(aa[i,j]);
                        }
                    }
                }
            }

            return re;
        }

        #endregion
        
    }
}
