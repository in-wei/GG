using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//-
using System.IO;
using System.Threading;
using UsuallyCommond.MyEnum;
using GlobalCommond;
using GlobalCommond.ViewModel;

namespace Report.Mrt
{
    public class MrtMain
    {

        #region 變數設定

        #region 設定(公開)

        public bool IsRunning { get; private set; }
        public List<MrtClient> MrtClients { get; private set; }
        public DateTime StartLoop { get; private set; }
        public DateTime OperationDate { get; set; }
        public DateTime OperationStartDate { get; set; }
        public DateTime OperationEndDate { get; set; }
        public string OperationLine { get; private set; }
        public GlobalCommonds globalCommond { get; set; }
        public List<MrtClient> StandByClient { get; private set; }
        public List<ReportList> List_ReportBox { get; private set; }

        public bool VA_Lock = false;

        public bool Od_get { get; private set; }
        public double Od_amt { get; private set; }

        public string TxnReportData_Str { get; private set; }
        public bool TxnReportData_Run { get; private set; }
        public int TxnReportData_Max { get; private set; }
        public int TxnReportData_Min { get; private set; }
        public int TxnReportData_Count { get; private set; }

        #endregion

        #region 交易資料

        public List<TxnTransaction> TxnAllData { get; private set; }
        public List<TxnTransaction> Data_External { get; private set; }
        public List<TxnTransactionOd> Data_Merge0 { get; private set; }
        public List<CardId> TxnAllCard { get; private set; }
        public List<CardId> TxnCard_more3 { get; private set; }
        public int CardCount_More { get; private set; }
        public List<CardId> TxnCard_less3 { get; private set; }
        public int CardCount_Less { get; private set; }
        public List<string> FinishCardSN { get; private set; }

        public List<TxnTransaction> ResponseExportTxnNormal { get; set; }
        public List<TxnTransaction> ResponseExportTxnCompare { get; set; }
        public List<TxnTransaction> ResponseExportTxnMeger { get; set; }
        public List<TxnTransaction> ResponseExportTxnError { get; set; }

        #endregion

        #region 比對數據

        //public List<ReportFileTemp2> Stock_Mrt_ReportFile2 { get; set; }
        //public List<MrtStationList_Old> Stock_Mrt_Station_Old { get; set; }
        //public List<MrtStationList> Stock_Mrt_Station { get; set; }
        //public List<MrtStationOd> Stock_Mrt_StationOd { get; set; }
        //public List<ReportFileTemp> Stock_Mrt_ReportFile { get; set; }
        //public List<MrtStationOdList> Stock_Mrt_OdList { get; set; }
        //public List<OdSubsidy> Stock_Mrt_Subsidy { get; set; }
        //public List<String_String_Int> Stock_IssureCode { get; set; }
        //public List<String_String_String_Int> Stock_Mrt_TxnType { get; set; }
        //public List<OnlyInt> Stock_Mrt_TxnType_Sub { get; set; }
        //public List<String_String_String_Int> Stock_TxnTypeSub_Ecc { get; set; }
        //public List<String_String_String_Int> Stock_TxnTypeSub_Ipass { get; set; }
        //public List<String_String_String_Int> Stock_TxnTypeSub_Icash { get; set; }
        //public List<String_String_String_Int> Stock_TxnTypeSub_Ticket { get; set; } //從300開始
        //public List<String_String_Int> Stock_FareProductType { get; set; }

        #endregion

        #region 檔案

        public string Config_Subidy = "Subsidy.csv";
        public string Config_Report = "Report.csv";
        public string Config_StationList_Old = "StationList_1.csv";
        public string Config_StationList = "StationList.csv";
        public string Config_StationOdList_Old = "MrtStationOdList_Old.csv";
        public string Config_StationOdList = "MrtStationOdList.csv";
        public string Config_IssureCode = "IssureCode.csv";
        public string Config_TxnType = "TxnType.csv";
        public string Config_TxnType_Sub = "TxnTypeSub.csv";
        public string Config_TxnTypeSub_Ecc = "TxnTypeSub_Ecc.csv";
        public string Config_TxnTypeSub_Ipass = "TxnTypeSub_Ipass.csv";
        public string Config_TxnTypeSub_Icash = "TxnTypeSub_Icash.csv";
        public string Config_Ticket_TxnTypeSub = "TxnTypeSub_Ticket.csv";
        public string Config_FareProductType = "FareProductType.csv";
        public string Switch_Operation_Line = "Mrt";

        #endregion

        #region 設定(封閉)

        private int CpuThread;
        private MyCommonds MyCommond;
        private string ThisReceive;
        bool StartStatus;
        bool FirstStart;
        private List<String_String_String_String_Int> StocketType;

        private string[] TxnType_In = { "20", "151", "153", "191", "194" };
        private string[] TxnType_Out = { "1", "2", "4", "8", "13", "152", "154", "192", "195" };

        Thread threadListen;
        Thread threadIsruning;
        Thread threadEzsyAnalyze;
        Thread threadTxnTypeAnalyze;
        Thread threadIOAnalyze;
        Thread threadCompareAnalyze;
        Thread threadTrnAnalyze;
        Thread threadVolumnReportAnalyze;
        Thread threadAccurateAnalyze;
        Thread threadC9Analyze;
        Thread threadC10Analyze;
        Thread threadC11Analyze;
        Thread threadC12Analyze;
        Thread threadC13Analyze;
        Thread threadC14Analyze;
        Thread threadC15Analyze;
        Thread threadC16Analyze;
        Thread threadC17Analyze;
        Thread threadC18Analyze;
        Thread threadC19Analyze;
        Thread threadC20Analyze;

        private Thread threadExportTxnNormal;
        private Thread threadExportTxnCompare;
        private Thread threadExportTxnMeger;
        private Thread threadExportTxnError;

        #endregion

        #endregion

        public MrtMain()
        {
            ThisReceive = "MainMrt";
            MyCommond = new MyCommonds();
            string ProgramFullPath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName.ToString();
            string ProgramLastWriteTime = File.GetLastWriteTime(ProgramFullPath).ToString();
            MyCommond.WriteLog(ThisReceive, $"This Program Last Write Time： {ProgramLastWriteTime}");
            TxnAllData = new List<TxnTransaction>();
            TxnAllCard = new List<CardId>();
            TxnCard_more3 = new List<CardId>();
            TxnCard_less3 = new List<CardId>();
            CardCount_More = 0;
            CardCount_Less = 0;
            IsRunning = false;
            FirstStart = true;
            StartStatus = false;
            //OperationLine = "環狀線";
            TxnReportData_Run = false;
            TxnReportData_Max = 0;
            TxnReportData_Min = 0;
            TxnReportData_Count = 0;
            //SubLoading subLoading = new SubLoading(this);

            if (!true) GetCpuThread();

            if (!true)
            {
                CompareInit();
            }

        }

        public void SetMyCommand()
        {
            MyCommond = globalCommond.MyCommond;
        }

        private void CompareInit()
        {

            ResponseExportTxnNormal = new List<TxnTransaction>();
            threadExportTxnError = new Thread(ExportTxnNormalFile);
            threadExportTxnError.IsBackground = true;
            threadExportTxnError.Start();

            ResponseExportTxnCompare = new List<TxnTransaction>();
            threadExportTxnCompare = new Thread(ExportTxnCompareFile);
            threadExportTxnCompare.IsBackground = true;
            threadExportTxnCompare.Start();

            ResponseExportTxnMeger = new List<TxnTransaction>();
            threadExportTxnMeger = new Thread(ExportTxnMegerFile);
            threadExportTxnMeger.IsBackground = true;
            threadExportTxnMeger.Start();

            ResponseExportTxnError = new List<TxnTransaction>();
            threadExportTxnError = new Thread(ExportTxnErrorFile);
            threadExportTxnError.IsBackground = true;
            threadExportTxnError.Start();

        }

        public MrtMain(DateTime OpenDateStart, DateTime OpenDateEnd) : this()
        {
            OperationStartDate = OpenDateStart;
            OperationEndDate = OpenDateEnd;
        }

        #region 功能

        public void Start(DateTime Dt1, DateTime Dt2)
        {
            






        }

        private void GetRuning()
        {
            try
            {
                while (IsRunning)
                {
                    while (waitLoop_2)
                    {
                        foreach (var item in MrtClients)
                        {
                            MyCommond.WriteLog(ThisReceive, $"{item.MrtMethodeName.ToDescription()} - Statue:{item.Status}");
                        }
                        Thread.Sleep(60000);
                    }
                }
            }
            catch (Exception ex)
            {
                //MyCommond.WriteLog(ThisReceive, $"ERROR!{ex.ToString()}");
            }
        }

        private void GetCpuThread()
        {
            CpuThread = Environment.ProcessorCount - 1;
            Console.WriteLine($"Processors {CpuThread}");
        }

        public bool waitLoop_1 = false;
        public bool waitLoop_2 = false;
        private Thread thread_ReWriteTxn;
        public void LoadTxn(CheckBoxStatus CB_Status, List<ReportList> tempReportList = null)
        {
            MyCommond.WriteLog(ThisReceive, $"執行運量開始");
            MrtClients = new List<MrtClient>();
            OperationStartDate = CB_Status.OperationStart;
            OperationEndDate = CB_Status.OperationEnd;
            bool waitTF = false;
            threadListen = new Thread(() =>
            {
                IsRunning = true;
                Thread threadStutes = new Thread(() => GetRuning());
                threadStutes.Start();
                /// List<Month_TxnTransaction> tea = new List<Month_TxnTransaction>();
                for (DateTime RoopDate = OperationStartDate; RoopDate < OperationEndDate; RoopDate = RoopDate.AddDays(1))
                {
                    waitLoop_1 = true;
                    OperationDate = RoopDate;
                    string downloadfilename = $"history{RoopDate:yyyyMMdd}_nc.csv";
                    string path = $@"{MyCommond.Path_Program}{MyCommond.Path_Analyze}{downloadfilename}";
                    MyCommond.WriteLog(ThisReceive, $"檢查檔案:{path}");
                    if (!MyCommond.CheckFile(path))
                    {
                        MyCommond.WriteLog(ThisReceive, $"未發現檔案");
                        continue;
                    }
                    Thread LoadTxnData = new Thread(() =>
                    {
                        MyCommond.WriteLog(ThisReceive, $"載入{RoopDate:yyyy/MM/dd}交易");
                        TxnAllData = MyCommond.ToViewModel<TxnTransaction>(ThisReceive, path, (new TxnTransaction().ExportTitle_English()), 1).ToList();
                        
                        if (!true)
                        {
                            if (LoadMrtTxnList(RoopDate)) { } // 繼續執行
                            else { goto CancelRun; }// 中斷執行
                        }

                        if (CB_Status.TrnAnalyze || CB_Status.VolumnReportAnalyze || CB_Status.EveryIssuerAnalyze) { ReWriteTxn(); }

                        StartLoop = DateTime.Now;

                        MyCommond.WriteLog(ThisReceive, $"選出與運量相關的交易類型");
                        StocketType = globalCommond.Stock_Mrt_TxnType.FindAll(x => x.mCheck == "EI" || x.mCheck == "EO" || x.mCheck == "TI" || x.mCheck == "TO").ToList();

                        /* 簡易資料(每日)        */ if (CB_Status.EasyAnalyze)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--EasyAnalyze--");
                            try
                            {
                                if (threadEzsyAnalyze.IsAlive) goto Next_1;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 簡易資料 站台");
                                threadEzsyAnalyze = new Thread(Start_Easy_Analyze);
                                threadEzsyAnalyze.IsBackground = true;
                                threadEzsyAnalyze.Start();
                            }
                        }
                    Next_1: /*  各交易資料       */ if (CB_Status.TxnTypeAnalyze)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--TxnTypeAnalyze--");
                            try
                            {
                                if (threadTxnTypeAnalyze.IsAlive) goto Next_2;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 各交易 站台");
                                threadTxnTypeAnalyze = new Thread(Start_TxnType_Analyze);
                                threadTxnTypeAnalyze.IsBackground = true;
                                threadTxnTypeAnalyze.Start();
                            }
                        }
                    Next_2: /*  進出站           */ if (CB_Status.IOAnalyze)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--IOAnalyze--");
                            try
                            {
                                if (threadIOAnalyze.IsAlive) goto Next_3;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 進出站 站台");
                                threadIOAnalyze = new Thread(Start_IO_Analyze);
                                threadIOAnalyze.IsBackground = true;
                                threadIOAnalyze.Start();
                            }
                        }
                    Next_3: /*  里程合併         */ if (CB_Status.CompareAnalyze)
                        {
                            //MyCommond.WriteLog(ThisReceive, $"篩選交易卡號");
                            MyCommond.WriteLog(ThisReceive, $"--CompareAnalyze--");
                            try
                            {
                                if (threadExportTxnNormal.IsAlive) goto Next_3_1;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                //MyCommond.WriteLog(ThisReceive, $"掛聽輸出線程(輸出不足3筆交易)");
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 延人里程合併 站台");
                                threadAccurateAnalyze = new Thread(Start_Compare_Analyze);
                                threadAccurateAnalyze.IsBackground = true;
                                threadAccurateAnalyze.Start();
                            }

                        Next_3_1:
                            MyCommond.WriteLog(ThisReceive, $"--ExportTxnNormal--");
                            try
                            {
                                if (threadCompareAnalyze.IsAlive) goto Next_3_2;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                CompareInit();
                                waitTF = true;
                            }

                        Next_3_2:
                            try
                            {

                            }
                            catch (Exception ex)
                            {
                                MyCommond.WriteLog(ThisReceive, $"Error! \n{ex.Message}");
                            }
                        }
                    Next_4: /*  轉乘             */ if (CB_Status.TrnAnalyze)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--TrnAnalyze--");
                            try
                            {
                                if (threadTrnAnalyze.IsAlive) goto Next_5;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 轉乘 站台");
                                threadTrnAnalyze = new Thread(Start_Trn_Analyze);
                                threadTrnAnalyze.IsBackground = true;
                                threadTrnAnalyze.Start();
                            }
                        }
                    Next_5: /*  通勤月票         */ if (CB_Status.TPassAnalyze)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--TPassAnalyze--");
                            try
                            {
                                if (threadIOAnalyze.IsAlive) goto Next_6;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 通勤月票 站台");
                                threadIOAnalyze = new Thread(Start_TPass_Analyze);
                                threadIOAnalyze.IsBackground = true;
                                threadIOAnalyze.Start();
                            }
                        }
                    Next_6: /*  運量報表         */ if (CB_Status.VolumnReportAnalyze)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--VolumnReportAnalyze--");
                            List_ReportBox = tempReportList;
                            try
                            {
                                if (threadVolumnReportAnalyze.IsAlive) goto Next_7;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 運量報表 站台");
                                threadVolumnReportAnalyze = new Thread(Start_VolumnReport_Analyze);
                                threadVolumnReportAnalyze.IsBackground = true;
                                threadVolumnReportAnalyze.Start();
                            }
                        }
                    Next_7: /*  加值報表         */ if (CB_Status.AddValueAnalyze)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--AddValueAnalyze--");
                            try
                            {
                                if (threadAccurateAnalyze.IsAlive) goto Next_8;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 AddValueAnalyze 站台");
                                threadAccurateAnalyze = new Thread(Start_AddValue_Analyze);
                                threadAccurateAnalyze.IsBackground = true;
                                threadAccurateAnalyze.Start();
                            }
                        }
                    Next_8: /*  各家支付運量日報 */ if (CB_Status.EveryIssuerAnalyze)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--EveryIssuerAnalyze--");
                            try
                            {
                                if (threadC9Analyze.IsAlive) goto Next_9;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 EveryIssuerAnalyze 站台");
                                threadC9Analyze = new Thread(Start_EveryIssuer_Analyze);
                                threadC9Analyze.IsBackground = true;
                                threadC9Analyze.Start();
                            }
                        }
                    Next_9: /*  check10          */ if (CB_Status.checkbox10)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--checkbox10--");
                            try
                            {
                                if (threadC10Analyze.IsAlive) goto Next_10;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 checkbox10 站台");
                                threadC10Analyze = new Thread(Start_C10_Analyze);
                                threadC10Analyze.IsBackground = true;
                                threadC10Analyze.Start();
                            }
                        }
                    Next_10: /* check11          */ if (CB_Status.checkbox11)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--checkbox11--");
                            try
                            {
                                if (threadC11Analyze.IsAlive) goto Next_11;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 checkbox11 站台");
                                threadC11Analyze = new Thread(Start_C11_Analyze);
                                threadC11Analyze.IsBackground = true;
                                threadC11Analyze.Start();
                            }
                        }
                    Next_11: /* check12          */ if (CB_Status.checkbox12)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--checkbox12--");
                            try
                            {
                                if (threadC12Analyze.IsAlive) goto Next_12;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 checkbox12 站台");
                                threadC12Analyze = new Thread(Start_C12_Analyze);
                                threadC12Analyze.IsBackground = true;
                                threadC12Analyze.Start();
                            }
                        }
                    Next_12: /* check13          */ if (CB_Status.checkbox13)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--checkbox13--");
                            try
                            {
                                if (threadC13Analyze.IsAlive) goto Next_13;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 checkbox13 站台");
                                threadC13Analyze = new Thread(Start_C13_Analyze);
                                threadC13Analyze.IsBackground = true;
                                threadC13Analyze.Start();
                            }
                        }
                    Next_13: /* check14          */ if (CB_Status.checkbox14)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--checkbox14--");
                            try
                            {
                                if (threadC14Analyze.IsAlive) goto Next_14;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 checkbox14 站台");
                                threadC14Analyze = new Thread(Start_C14_Analyze);
                                threadC14Analyze.IsBackground = true;
                                threadC14Analyze.Start();
                            }
                        }
                    Next_14: /* check15          */ if (CB_Status.checkbox15)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--checkbox15--");
                            try
                            {
                                if (threadC15Analyze.IsAlive) goto Next_15;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 checkbox15 站台");
                                threadC15Analyze = new Thread(Start_C15_Analyze);
                                threadC15Analyze.IsBackground = true;
                                threadC15Analyze.Start();
                            }
                        }
                    Next_15: /* check16          */ if (CB_Status.checkbox16)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--checkbox16--");
                            try
                            {
                                if (threadC16Analyze.IsAlive) goto Next_16;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 checkbox16 站台");
                                threadC16Analyze = new Thread(Start_C16_Analyze);
                                threadC16Analyze.IsBackground = true;
                                threadC16Analyze.Start();
                            }
                        }
                    Next_16: /* check17          */ if (CB_Status.checkbox17)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--checkbox17--");
                            try
                            {
                                if (threadC17Analyze.IsAlive) goto Next_17;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 checkbox17 站台");
                                threadC17Analyze = new Thread(Start_C13_Analyze);
                                threadC17Analyze.IsBackground = true;
                                threadC17Analyze.Start();
                            }
                        }
                    Next_17: /* check18          */ if (CB_Status.checkbox18)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--checkbox18--");
                            try
                            {
                                if (threadC18Analyze.IsAlive) goto Next_18;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 checkbox18 站台");
                                threadC18Analyze = new Thread(Start_C13_Analyze);
                                threadC18Analyze.IsBackground = true;
                                threadC18Analyze.Start();
                            }
                        }
                    Next_18: /* check19          */ if (CB_Status.checkbox19)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--checkbox19--");
                            try
                            {
                                if (threadC19Analyze.IsAlive) goto Next_19;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 checkbox19 站台");
                                threadC19Analyze = new Thread(Start_C19_Analyze);
                                threadC19Analyze.IsBackground = true;
                                threadC19Analyze.Start();
                            }
                        }
                    Next_19: /* check20          */ if (CB_Status.checkbox20)
                        {
                            MyCommond.WriteLog(ThisReceive, $"--checkbox20--");
                            try
                            {
                                //if (threadC20Analyze.IsAlive) goto Next_20;
                            }
                            catch (Exception ex)
                            {

                            }
                            finally
                            {
                                MyCommond.WriteLog(ThisReceive, $"掛聽啟動 checkbox20 站台");
                                threadC20Analyze = new Thread(Start_C20_Analyze);
                                threadC20Analyze.IsBackground = true;
                                threadC20Analyze.Start();
                            }
                        }

                    waitThread:
                        if (CB_Status.TrnAnalyze || CB_Status.VolumnReportAnalyze || CB_Status.EveryIssuerAnalyze) { thread_ReWriteTxn.Join(); }
                        waitLoop_2 = true;
                        goto EndThread;

                    CancelRun:
                        waitLoop_1 = false;
                        waitLoop_2 = false;
                        IsRunning = false;
                        goto EndThread;

                    jumpThread:
                        waitLoop_1 = false;
                        waitLoop_2 = false;
                        goto EndThread;

                    EndThread:
                        int i = 0;
                    });
                    LoadTxnData.Start();
                    //Thread.Sleep(30000);
                    while (waitLoop_1)
                    {
                        while (waitLoop_2)
                        {
                            try
                            {
                                Thread.Sleep(5000);
                                if (MrtClients.Count == 0) { waitLoop_1 = false; waitLoop_2 = false; }
                                var ListFindFinish = MrtClients.FindAll(x => x.Status == StatusEnum.Finish).ToList();
                                var ListDel = MrtClients.FindAll(x => x.Status != StatusEnum.Calculating).ToList();
                                if (ListFindFinish.Count == 0) continue;
                                if (MrtClients.Count() != ListDel.Count()) continue;
                                foreach (var item in ListDel) { MrtClients.Remove(item); }
                                //MrtClients.Clear();
                                //MrtClients = null;
                                TxnAllData.Clear();


                                waitLoop_1 = false;
                                waitLoop_2 = false;
                            }
                            catch (Exception ex)
                            {
                                MyCommond.WriteLog(ThisReceive, $"ERROR!\t{ex.Message}");
                            }
                        }
                    }
                    while (waitTF)
                    {
                        if ( waitTF) waitTF = ResponseExportTxnCompare.Count != 0;
                        if (!waitTF) waitTF = ResponseExportTxnError.Count != 0;
                        if (!waitTF) waitTF = ResponseExportTxnMeger.Count != 0;
                        if (!waitTF) waitTF = ResponseExportTxnNormal.Count != 0;
                    }
                    try
                    {
                        MyCommond.MoveFile($"{MyCommond.Path_Program}{MyCommond.Path_Analyze}",
                            $@"{MyCommond.Path_Program}{MyCommond.Path_Analyze}Old\",
                            $"history{RoopDate:yyyyMMdd}_nc.csv");
                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $"MoveFileError!Msg:{ex.Message}");
                    }
                }
                
                /// foreach (var item in tea)
                /// {
                ///     MyCommond.WriteLog(ThisReceive, item.ToString());
                ///     Console.WriteLine(item.ToString());
                /// }

                for (DateTime rootTime = OperationStartDate; rootTime < OperationEndDate; rootTime = rootTime.AddDays(1))
                {
                    try
                    {
                        MyCommond.MoveFile($"{MyCommond.Path_Program}{MyCommond.Path_Analyze}",
                            $@"{MyCommond.Path_Program}{MyCommond.Path_Analyze}Old\",
                            $"history{rootTime:yyyyMMdd}_nc.csv");
                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $"MoveFileError!Msg:{ex.Message}");
                    }
                }
                
                IsRunning = false;
            });
            threadListen.IsBackground = true;
            threadListen.Start();
        }

        public void Stop()
        {
            StartStatus = false;
            IsRunning = false;
            lock (MrtClients)
            {
                List<MrtClient> clients = MrtClients.ToList();
                foreach (var item in clients)
                {
                    MrtClients.Remove(item);
                }
            }
        }

        public string GetCardSN(string GetID)
        {
            var returnMsg = TxnCard_more3[CardCount_More++].CardSN;
            Console.WriteLine($"{GetID} 取得卡號 {returnMsg,20}");
            //MyCommond.WriteLog(ThisReceive, ExecutionMode.Simple, $"{GetID} 取得卡號 {returnMsg,20} ");
            return returnMsg;
        }

        public List<TxnTransaction> GetCardTxnData(string GetID, string CardSN)
        {
            List<TxnTransaction> result = new List<TxnTransaction>();
            result = (from p in TxnAllData where p.CARD_PHYSICAL_ID == CardSN orderby p.TXN_TIMESTAMP select p).ToList();
            MyCommond.WriteLog(ThisReceive, $"No.{CardCount_More + CardCount_Less},ID:{GetID} 取得卡號 {CardSN,20} 交易資料 {result.Count()} 筆");
            return result;
        }

        private void ReWriteTxn()
        {
            TxnReportData_Min = 0;
            TxnReportData_Max = 100;
            TxnReportData_Count = 0;
            TxnReportData_Str = "";
            TxnReportData_Run = true;

            List<int> ls_O = new List<int>();
            string ls_O_str = "";
            string ReportPath = MyCommond.Path_Program + MyCommond.Path_Report + @"\環狀線\" + string.Format(MyCommond.Path_Date, OperationDate.ToString("yyyy.MM.dd"));
            MyCommond.WriteLog(ThisReceive, $"檔案路徑2:{ReportPath}");
            var station_Od = globalCommond.Stock_Mrt_Y_OdList.ToList();
            var station1 = globalCommond.Stock_Mrt_Y_Station;
            
            #region 取出比對樣本

            foreach (var item in globalCommond.Stock_Mrt_TxnType)
            {
                if (item.mCheck != null)
                {
                    if (item.mCheck.Length == 2 && item.mCheck.Substring(1, 1) == "O")
                    {
                        ls_O.Add(item.Num);
                        ls_O_str += $"{item.Num},";
                    }
                    else if (item.mCheck.Length == 1 && item.mCheck == "T")
                    {
                        ls_O.Add(item.Num);
                        ls_O_str += $"{item.Num},";
                    }
                }
            }
            MyCommond.WriteLog(ThisReceive, $"出站變數共 {ls_O.Count} 項，{ls_O_str.Substring(0, ls_O_str.Length - 1)}");

            List<string> station2 = (from p in station1 select p.CodeName1).ToList();
            MyCommond.WriteLog(ThisReceive, $"取出全車站代號共 {station2.Count} 項");
            //List<string> stationTransfer = (from x in station_ where x.CodeName == "82" || x.CodeName == "83" || x.CodeName == "209" || x.CodeName == "210" select x.StationName).ToList();
            string[] sss = new string[] { "板橋", "新埔" };
            List<string> stationTransfer2 = new List<string>();
            foreach (var item in station1)
            {
                foreach (var jtem in sss)
                {
                    if (System.Text.RegularExpressions.Regex.IsMatch(item.StationName1, jtem))
                    {
                        stationTransfer2.Add(item.StationName1);
                    }
                }
            }
            MyCommond.WriteLog(ThisReceive, $"取出站外轉乘車站代號      共 {stationTransfer2.Count,8} 項");

            var Stock_Ecc /*---*/ = (from p in globalCommond.Stock_Mrt_TxnTypeSub_Ecc /*---*/ where p.mCheck == "o" select p.Num).ToList();
            MyCommond.WriteLog(ThisReceive, $"取出     悠遊卡子     類別共 {Stock_Ecc.Count,8} 項");
            var Stock_Ipass /*-*/ = (from p in globalCommond.Stock_Mrt_TxnTypeSub_Ipass /*-*/ where p.mCheck == "o" select p.Num).ToList();
            MyCommond.WriteLog(ThisReceive, $"取出     一卡通子     類別共 {Stock_Ipass.Count,8} 項");
            var Stock_Icash /*-*/ = (from p in globalCommond.Stock_Mrt_TxnTypeSub_Icash /*-*/ where p.mCheck == "o" select p.Num).ToList();
            MyCommond.WriteLog(ThisReceive, $"取出     愛金卡子     類別共 {Stock_Icash.Count,8} 項");

            #endregion

            #region 運量資料處理 1
            var Data_Unmodified = (from p in TxnAllData where ls_O.Contains(Convert.ToInt32(p.CARD_TXN_TYPE_ID)) /*&& station.Contains(p.ENTRY_LOC_ID) && station.Contains(p.SVCE_LOC_ID)*/ select p).ToList();
            MyCommond.WriteLog(ThisReceive, $"取出       出站       資料共 {Data_Unmodified.Count,8} 筆");
            List<TxnTransaction> Data_Ecc___ = (from p in Data_Unmodified where /*---*/ Stock_Ecc.Contains(Convert.ToInt32(p.CARD_TXN_SUBTYPE_ID)) && Convert.ToInt32(p.CARD_TXN_TYPE_ID) <= 150 && p.ISSUER_ID == "255" /*---------*/ select p).ToList();
            MyCommond.WriteLog(ThisReceive, $"取出      悠遊卡      資料共 {Data_Ecc___.Count,8} 筆");
            List<TxnTransaction> Data_Ipass_ = (from p in Data_Unmodified where /*-*/ Stock_Ipass.Contains(Convert.ToInt32(p.CARD_TXN_SUBTYPE_ID)) && Convert.ToInt32(p.CARD_TXN_TYPE_ID) <= 150 && p.ISSUER_ID == "9" /*-----------*/ select p).ToList();
            MyCommond.WriteLog(ThisReceive, $"取出      一卡通      資料共 {Data_Ipass_.Count,8} 筆");
            List<TxnTransaction> Data_Icash_ = (from p in Data_Unmodified where /*-*/ Stock_Icash.Contains(Convert.ToInt32(p.CARD_TXN_SUBTYPE_ID)) && Convert.ToInt32(p.CARD_TXN_TYPE_ID) <= 150 && p.ISSUER_ID == "11" /*----------*/ select p).ToList();
            MyCommond.WriteLog(ThisReceive, $"取出      愛金卡      資料共 {Data_Icash_.Count,8} 筆");
            List<TxnTransaction> Data_Mobile = (from p in Data_Unmodified where /*---*/ Convert.ToInt32(p.CARD_TXN_TYPE_ID) <= 150 && 20 <= Convert.ToInt16(p.ISSUER_ID) && Convert.ToInt16(p.ISSUER_ID) < 100 select p).ToList();
            MyCommond.WriteLog(ThisReceive, $"取出     行動支付     資料共 {Data_Mobile.Count,8} 筆");
            List<TxnTransaction> Data_Ticket = (from p in Data_Unmodified where /*-----------------------------------------------------------------*/ Convert.ToInt32(p.CARD_TXN_TYPE_ID) > 150 /*--------------------------*/ select p).ToList();
            MyCommond.WriteLog(ThisReceive, $"取出      自有票      資料共 {Data_Ticket.Count,8} 筆");

            bool EachDataExport = false;
            if (EachDataExport)
            {
                MyCommond.WriteLog(ThisReceive, $"各票種輸出");
                string VerifyPath = MyCommond.Path_Program + MyCommond.Path_Verify;
                MyCommond.WriteLog(ThisReceive, $"悠遊卡輸出");
                MyCommond.ExportData(ThisReceive, new ExportConfig()
                {
                    Path1 = VerifyPath,
                    DTime = OperationDate,
                    FirstName = "",
                    FileName = "DataEcc",
                    SubName = "",
                    OpLine = "環狀線",
                    LastName = "csv",

                }, new TxnTransaction().ExportTitle_zhTW(), Data_Ecc___, Encoding.UTF8);
                MyCommond.WriteLog(ThisReceive, $"一卡通輸出");
                MyCommond.ExportData(ThisReceive, new ExportConfig()
                {
                    Path1 = VerifyPath,
                    DTime = OperationDate,
                    FirstName = "",
                    FileName = "DataIpass",
                    SubName = "",
                    OpLine = "環狀線",
                    LastName = "csv",

                }, new TxnTransaction().ExportTitle_zhTW(), Data_Icash_, Encoding.UTF8);
                MyCommond.WriteLog(ThisReceive, $"愛金卡輸出");
                MyCommond.ExportData(ThisReceive, new ExportConfig()
                {
                    Path1 = VerifyPath,
                    DTime = OperationDate,
                    FirstName = "",
                    FileName = "DataIcash",
                    SubName = "",
                    OpLine = "環狀線",
                    LastName = "csv",

                }, new TxnTransaction().ExportTitle_zhTW(), Data_Icash_, Encoding.UTF8);
                MyCommond.WriteLog(ThisReceive, $"行動支付輸出");
                MyCommond.ExportData(ThisReceive, new ExportConfig()
                {
                    Path1 = VerifyPath,
                    DTime = OperationDate,
                    FirstName = "",
                    FileName = "DataMobile",
                    SubName = "",
                    OpLine = "環狀線",
                    LastName = "csv",

                }, new TxnTransaction().ExportTitle_zhTW(), Data_Mobile, Encoding.UTF8);
                MyCommond.WriteLog(ThisReceive, $"自有票輸出");
                MyCommond.ExportData(ThisReceive, new ExportConfig()
                {
                    Path1 = VerifyPath,
                    DTime = OperationDate,
                    FirstName = "",
                    FileName = "DataTicket",
                    SubName = "",
                    OpLine = "環狀線",
                    LastName = "csv",

                }, new TxnTransaction().ExportTitle_zhTW(), Data_Ticket, Encoding.UTF8);
                int wait = 0;
            }

            MyCommond.WriteLog(ThisReceive, $"將所有票種合併");
            List<TxnTransaction> Data_Merge = new List<TxnTransaction>();
            Data_Merge.AddRange(Data_Ecc___);
            Data_Merge.AddRange(Data_Ipass_);
            Data_Merge.AddRange(Data_Icash_);
            Data_Merge.AddRange(Data_Mobile);
            Data_Merge.AddRange(Data_Ticket);
            TxnReportData_Max = Data_Merge.Count;
            #endregion 運量資料處理 1
            #region 運量資料處理 2
        分析:
            MyCommond.WriteLog(ThisReceive, $"{(TxnReportData_Str = "1 / 6 - 將合併的資料重新編輯")}");
            List<TxnTransactionOd> Data_Merge2 = new List<TxnTransactionOd>();

            double _shoudPay = -1, _realPay = -1, _allPassAmount = -1, _discAmount = -1, _discPoint = -1, _trnNormal = -1, _trnAllPass = -1, _off_10 = -1;
            double _splitShoudPay = -1, _splitRealPay = -1, _splitAllPassAmount = -1, _splitDiscAmount = -1, _splitDiscPoint = -1, _splitTrnNormal = -1, _splitTrnAllPass = -1, _splitOff_10 = -1;
            double _distanceRatio = -1, _amountMark = -1;
            string _txntype = "", _txnsub = "", _issuer = "", _entry = "", _exit = "", _ticketType = "", _errorString = "", _ticket = "", _fare = "";
            UsuallyCommond.MyEnum.Ele_Iss _issuer2 = UsuallyCommond.MyEnum.Ele_Iss.NA;

            Data_Merge2 = (from p in Data_Merge orderby p.TXN_TIMESTAMP select new TxnTransactionOd
            {
                #region 普通
                /*--------*/ CARD_TXN_TYPE_ID = Check_TxnType(p.CARD_TXN_TYPE_ID, ref _txntype),
                /*-----*/ CARD_TXN_SUBTYPE_ID = Check_TxnSubType(Check_IssuerId(p.ISSUER_ID, p.CARD_TXN_TYPE_ID, ref _issuer, ref _issuer2), p.CARD_TXN_SUBTYPE_ID, ref _txnsub),
                /*-----*/ //CARD_TXN_SUBTYPE_ID = Check_TxnSubType((p.ISSUER_ID == "255") ? (Ele_Iss.ECC) : ((p.ISSUER_ID == "9") ? (Ele_Iss.Ipass) : ((p.ISSUER_ID == "11") ? (Ele_Iss.Icash) : (Ele_Iss.NA))), p.CARD_TXN_SUBTYPE_ID, ref _txnsub),
                /*------------------*/ DEV_ID = p.DEV_ID,
                /*-----------*/ TXN_TIMESTAMP = p.TXN_TIMESTAMP,
                /*--------*/ CARD_PHYSICAL_ID = Check_CardId(p.CARD_PHYSICAL_ID, p.ISSUER_ID, p.SAM_ID),
                /*---------------*/ ISSUER_ID = _issuer,
                /*---------------*/ //ISSUER_ID = Check_IssuerId(p.ISSUER_ID, p.CARD_TXN_TYPE_ID, ref _issuer),
                /*---------*/ CARD_TXN_SEQ_NO = p.CARD_TXN_SEQ_NO,
                /*-----------------*/ TXN_AMT = p.TXN_AMT,
                /*--------*/ ELECTRONIC_VALUE = p.ELECTRONIC_VALUE,
                /*-------------*/ SVCE_LOC_ID = Check_Station(p.SVCE_LOC_ID, ref _exit),
                /*---------*/ PROCESSING_DATE = p.PROCESSING_DATE,
                /*-----------*/ BUSINESS_DATE = p.BUSINESS_DATE,
                /*------------*/ ENTRY_LOC_ID = Check_Station(p.ENTRY_LOC_ID, ref _entry),
                /*---------------*/ XFER_CODE = p.XFER_CODE,
                /*---------------*/ XFER_DISC = p.XFER_DISC,
                /*-----------*/ PERSONAL_DISC = p.PERSONAL_DISC,
                /*-----------------*/ PENALTY = p.PENALTY,
                /*---------*/ LOYALTY_COUNTER = p.LOYALTY_COUNTER,
                /*----------*/ LOYALTY_POINTS = p.LOYALTY_POINTS,
                /*----*/ FARE_PRODUCT_TYPE_ID = Check_Fare(p.FARE_PRODUCT_TYPE_ID, ref _fare),
                /*----------*/ ENTRY_DATETIME = p.ENTRY_DATETIME,
                /*---------------*/ AREA_CODE = p.AREA_CODE,
                /*-------*/ TICKET_SUBTYPE_ID = Check_Ticket(p.TICKET_SUBTYPE_ID, ref _ticket),
                /*------*/ XFER_DISC_BUSTOMRT = p.XFER_DISC_BUSTOMRT,
                /*------------*/ USER_PROFILE = p.USER_PROFILE,
                /*--*/ FIRST_UTILISATION_DATE = p.FIRST_UTILISATION_DATE,
                /*-----*/ UP_UTILISATION_DATE = p.UP_UTILISATION_DATE,
                /*---*/ LAST_UTILISATION_DATE = p.LAST_UTILISATION_DATE,
                /*-------------------*/SAM_ID = p.SAM_ID,
                /*-----------------*/ M_Entry = $"{Check_IO(station_Od, E2.Entry, _entry, _exit)}",
                /*------------------*/ M_Exit = $"{Check_IO(station_Od, E2.Exit, _entry, _exit)}",
                /*--------------*/ TicketType = $"{Check_TicketType(_issuer2, p.CARD_TXN_TYPE_ID, p.CARD_TXN_SUBTYPE_ID, p.FARE_PRODUCT_TYPE_ID, p.TICKET_SUBTYPE_ID, ref _ticketType)}",
                /*------------*/ TransferMark = $"{Check_TransferMark(stationTransfer2, _entry, _exit)}",
                /*------------------*/ Remark = $"{Check_Remark(_entry, _exit)}",
                #endregion 普通
                #region 新增
                /*----------------*/ ShoudPay = $"{ShoudPay(p.TXN_AMT, _entry, _exit, ref _shoudPay, ref _errorString)}", 
                /*-----------*/ AllPassAmount = $"{AllPassAmount(p.TXN_AMT, _ticketType, ref _allPassAmount, ref _errorString)}",
                /*--------------*/ DiscAmount = $"{DiscAmount(_shoudPay, p.PERSONAL_DISC, p.AREA_CODE, _ticketType, ref _discAmount, ref _errorString)}",
                /*---------------*/ DiscPoint = $"{DiscPoin(p.LOYALTY_COUNTER, _ticketType, _discAmount, ref _discPoint, ref _errorString)}",
                /*---------*/ TrnNormalAmount = $"{TrnNormal(p.XFER_DISC, _allPassAmount, ref _trnNormal, ref _errorString)}",
                /*--------*/ TrnAllPassAmount = $"{TrnAllPass(p.XFER_DISC, _allPassAmount, ref _trnAllPass, ref _errorString)}",
                /*------------------*/ Off_10 = $"{Off_10(_shoudPay, _discAmount, ref _off_10, ref _errorString)}",
                /*-----------------*/ RealPay = $"{RealPay(p.TXN_AMT, p.LOYALTY_COUNTER, _shoudPay, _allPassAmount, _discAmount, _off_10, ref _realPay, ref _errorString)}",
                
                /*-----------*/ DistanceRatio = $"{DistanceRatio(_entry, _exit, ref _distanceRatio, ref _errorString)}",

                /*----------*/ Split_ShoudPay = $"{Split_ShoudPay(_shoudPay, _distanceRatio, ref _errorString)}",
                /*-----------*/ Split_RealPay = $"{Split_RealPay(_realPay, _distanceRatio, ref _errorString)}",
                /*-----*/ Split_AllPassAmount = $"{Split_AllPassAmount(_allPassAmount, _distanceRatio, ref _errorString)}",
                /*--------*/ Split_DiscAmount = $"{Split_DiscAmount(_discAmount, _distanceRatio, ref _errorString)}",
                /*---------*/ Split_DiscPoint = $"{Split_DiscPoint(_discPoint, _distanceRatio, ref _errorString)}",
                /*---*/ Split_TrnNormalAmount = $"{Split_TrnNormal(_trnNormal, _distanceRatio, ref _errorString)}",
                /*--*/ Split_TrnAllPassAmount = $"{Split_TrnAllPass(_trnAllPass, _distanceRatio, ref _errorString)}",
                /*----------- */ Split_Off_10 = $"{Split_Off_10(_off_10, _distanceRatio, ref _errorString)}",

                /*-----------*/ TicketSubsidy = ticketSidy(_entry, _exit, ref _errorString),

                /*------------*/ AmountRemark = $"{_errorString}",

                #endregion 新增
            }).ToList();

            #region a

            MyCommond.WriteLog(ThisReceive, $"全資料              總數量共 {Data_Merge2.Count,8} 筆");

            MyCommond.WriteLog(ThisReceive, $"{(TxnReportData_Str = "2 / 6 - 篩選出本司運量")}");
            List<TxnTransactionOd> Data_Merge3 = (from p in Data_Merge2 where station_Od.FindIndex(x => x.IsPass == "Y" && x.Entry == p.ENTRY_LOC_ID && x.Exit == p.SVCE_LOC_ID) > -1 select p).ToList();
            MyCommond.WriteLog(ThisReceive, $"本司的              總數量共 {Data_Merge3.Count,8} 筆");

            MyCommond.WriteLog(ThisReceive, $"{(TxnReportData_Str = "3 / 6 - 篩選出非本司運量或Null")}");
            List<TxnTransactionOd> Data_Merge4 = (from p in Data_Merge2 where !(station_Od.FindIndex(x => x.IsPass == "Y" && x.Entry == p.ENTRY_LOC_ID && x.Exit == p.SVCE_LOC_ID) > -1) select p).ToList();
            MyCommond.WriteLog(ThisReceive, $"非本司              總數量共 {Data_Merge4.Count,8} 筆");

            MyCommond.WriteLog(ThisReceive, $"{(TxnReportData_Str = "4 / 6 - 篩選出本司的NULL進新北出")}");
            List<TxnTransactionOd> Data_Merge5 = (from p in Data_Merge4 where p.ENTRY_LOC_ID == "NULL" select p).ToList();
            MyCommond.WriteLog(ThisReceive, $"本司的NULL進新北出  總數量共 {Data_Merge5.Count,8} 筆");

            MyCommond.WriteLog(ThisReceive, $"{(TxnReportData_Str = "5 / 6 - 篩選出非本司的進出站或NULL")}");
            List<TxnTransactionOd> Data_Merge6 = (from p in Data_Merge4 where !(p.ENTRY_LOC_ID == "NULL") select p).ToList();
            MyCommond.WriteLog(ThisReceive, $"非本司的進出站或NULL總數量共 {Data_Merge6.Count,8} 筆");

            MyCommond.WriteLog(ThisReceive, $"{(TxnReportData_Str = "6 / 6 - 合併有關本司的資料")}");
            Data_Merge0 = new List<TxnTransactionOd>();
            Data_Merge0.AddRange(Data_Merge3);
            Data_Merge0.AddRange(Data_Merge5);

            #endregion a

            #endregion 運量資料處理 2

            Data_Merge0 = (from p in Data_Merge0 orderby p.TXN_TIMESTAMP select p).ToList();
            MyCommond.WriteLog(ThisReceive, $"有關本司的資料      總數量共 {Data_Merge0.Count,8} 筆");
            bool isExport = false;
            if (isExport)
            {
                MyCommond.WriteLog(ThisReceive, $"匯出各類型交易資料輸出");
                MyCommond.WriteLog(ThisReceive, $"匯出全運量資料輸出"); /*-------*/ MyCommond.ExportData(ThisReceive, new ExportConfig()
                {
                    Path1 = ReportPath,
                    DTime = OperationDate,
                    FirstName = "",
                    FileName = "DataTxnOd",
                    SubName = "All_1",
                    OpLine = "環狀線",
                    LastName = "xlsx",
                }, new TxnTransaction().ExportTitle_zhTW(), Data_Merge, Encoding.UTF8);
                MyCommond.WriteLog(ThisReceive, $"匯出全運量資料重編輸出"); /*---*/ MyCommond.ExportData(ThisReceive, new ExportConfig()
                {
                    Path1 = ReportPath,
                    DTime = OperationDate,
                    FirstName = "",
                    FileName = "DataTxnOd",
                    SubName = "All_2",
                    OpLine = "環狀線",
                    LastName = "xlsx",
                }, new TxnTransactionOd().ExportTitle_zhTW(), Data_Merge2, Encoding.UTF8);
                MyCommond.WriteLog(ThisReceive, $"匯出本司運量資料輸出"); /*-----*/ MyCommond.ExportData(ThisReceive, new ExportConfig()
                {
                    Path1 = ReportPath,
                    DTime = OperationDate,
                    FirstName = "",
                    FileName = "DataTxnOd",
                    SubName = "OnLine",
                    OpLine = "環狀線",
                    LastName = "xlsx",
                }, new TxnTransactionOd().ExportTitle_zhTW(), Data_Merge3, Encoding.UTF8);
                MyCommond.WriteLog(ThisReceive, $"匯出非本司運量或Null"); /*-----*/ MyCommond.ExportData(ThisReceive, new ExportConfig()
                {
                    Path1 = ReportPath,
                    DTime = OperationDate,
                    FirstName = "",
                    FileName = "DataTxnOd",
                    SubName = "AllNot",
                    OpLine = "環狀線",
                    LastName = "xlsx",
                }, new TxnTransactionOd().ExportTitle_zhTW(), Data_Merge4, Encoding.UTF8);
                MyCommond.WriteLog(ThisReceive, $"匯出本司的NULL進新北出"); /*---*/ MyCommond.ExportData(ThisReceive, new ExportConfig()
                {
                    Path1 = ReportPath,
                    DTime = OperationDate,
                    FirstName = "",
                    FileName = "DataTxnOd",
                    SubName = "Null",
                    OpLine = "環狀線",
                    LastName = "xlsx",
                }, new TxnTransactionOd().ExportTitle_zhTW(), Data_Merge5, Encoding.UTF8);
                MyCommond.WriteLog(ThisReceive, $"匯出非本司的進出站或NULL"); /*-*/ MyCommond.ExportData(ThisReceive, new ExportConfig()
                {
                    Path1 = ReportPath,
                    DTime = OperationDate,
                    FirstName = "",
                    FileName = "DataTxnOd",
                    SubName = "Not",
                    OpLine = "環狀線",
                    LastName = "xlsx",
                }, new TxnTransactionOd().ExportTitle_zhTW(), Data_Merge6, Encoding.UTF8);
                int wait = 0;
            }

            thread_ReWriteTxn = new Thread(() =>
            {
                MyCommond.WriteLog(ThisReceive, $"匯出有關本司的資料 - 開始");
                MyCommond.ExportData(ThisReceive, new ExportConfig()
                {
                    Path1 = ReportPath,
                    DTime = OperationDate,
                    FirstName = "",
                    FileName = "運量交易資料",
                    SubName = "",
                    OpLine = "環狀線",
                    LastName = "xlsx",
                }, new TxnTransactionOd().ExportTitle_zhTW(), Data_Merge0, Encoding.UTF8);
                MyCommond.WriteLog(ThisReceive, $"匯出有關本司的資料 - 結束");
            });
            thread_ReWriteTxn.IsBackground = true;
            if ( true) thread_ReWriteTxn.Start();
            //thread_E.Join();
            TxnReportData_Run = false;
        }

        private bool LoadMrtTxn()
        {
            string Download_FilePath = $"{MyCommond.Path_Analyze}";
            string Download_FileName = $"history{OperationDate.ToString("yyyyMMdd")}_nc";
            string Download_LastName = "csv";
            string Download_FullPath = Download_FilePath + Download_FileName + "." + Download_LastName;

            string OtherDate_Path = $"./OtherDate/";
            string OtherDate_FileName = "history{0}_cut";
            string OtherDate_LastName = "csv";
            string OtherDateFullPath = OtherDate_Path + OtherDate_FileName + "." + OtherDate_LastName;

            string CheckWrite_Path = $"./OtherDate/";
            string CheckWrite_FileName = $"切分檔案確認";
            string CheckWrite_LastName = "xlsx";
            string CheckWrite_FullPath = CheckWrite_Path + CheckWrite_FileName + "." + CheckWrite_LastName;

            var DownloadData_1 = MyCommond.ToViewModel<TxnTransaction>(ThisReceive, Download_FullPath, (new TxnTransaction().ExportTitle_English()), 1).ToList();

            bool checkfile_checkwrite = MyCommond.CheckFile(CheckWrite_FullPath);
            List<LoadData> CheckWriteData_1 = MyCommond.ToViewModel<LoadData>(ThisReceive, CheckWrite_FullPath, (new LoadData().ExportTitle_English()), 1).ToList();
            LoadData ExportloadData = new LoadData()
            {
                No = (checkfile_checkwrite) ? CheckWriteData_1.Count.ToString() : "1",
                FileName = Download_FileName + "." + Download_LastName,
                InstallDate = OperationDate.ToString("yyyy/MM/dd"),
                StartDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"),
                FinishDate = "",
            };

            





            return false;
        }

        private bool LoadMrtTxnList(DateTime LoadOperationDate)
        {
            string LoadDataTitle = new LoadData().ExportTitle_zhTW();

            string CheckWrite_Path = $@"{MyCommond.Path_Program}OtherDate\";
            string CheckWrite_FileName = $"切分檔案確認";
            FileType CheckWrite_LastName = FileType.CSV;
            string CheckWrite_FullPath = $"{CheckWrite_Path}{LoadOperationDate.ToString("yyyyMM")}_{CheckWrite_FileName}.{CheckWrite_LastName.ToDescription()}";
            bool checkfile_checkwrite = MyCommond.CheckFile(CheckWrite_FullPath);
            if (!checkfile_checkwrite) { ReWriteData<LoadData>(CheckWrite_FullPath, LoadDataTitle, null, FileType.CSV, Encoding.UTF8); }

            string Download_FilePath = $"{MyCommond.Path_Program}{MyCommond.Path_Analyze}";
            string Download_FileName = $"history{OperationDate.ToString("yyyyMMdd")}_nc";
            FileType Download_LastName = FileType.CSV;
            string Download_FullPath = $"{Download_FilePath}{Download_FileName}.{Download_LastName.ToDescription()}";

            int CheckWriteData_1 = File.ReadAllLines(CheckWrite_FullPath, Encoding.Default).Length - 1;

            LoadData ExportloadData = new LoadData()
            {
                No = (checkfile_checkwrite) ? CheckWriteData_1.ToString() : "1",
                FileName = $"{Download_FileName}.{Download_LastName}",
                InstallDate = OperationDate.ToString("yyyy/MM/dd"),
                StartDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"),
                FinishDate = "",
            };
            
            bool uploadData_TF = true;
            uploadData_TF = !UpLoaddata(CheckWrite_FullPath, LoadDataTitle, ExportloadData, CheckWrite_LastName); // 第一次寫入資料
            if (uploadData_TF) return false;  //發現已經寫入過了
            MyCommond.WriteLog(ThisReceive, $"開始");

            Thread.Sleep(10000);

            string OtherDate_Path = $@"{MyCommond.Path_Program}\OtherDate\";
            string OtherDate_FileName = "history{0}_cut";
            FileType OtherDate_LastName = FileType.CSV;
            string OtherDateFullPath = $"{OtherDate_Path}{OtherDate_FileName}.{OtherDate_LastName.ToDescription()}";

            var LoadTxnBusiness = (from p in TxnAllData group p by new { p.BUSINESS_DATE } into g select g.Key.BUSINESS_DATE).ToList();
            
            foreach (var item in LoadTxnBusiness)
            {
                var itemTxn = (from p in TxnAllData where p.BUSINESS_DATE == item select p).ToList();

                if (true)
                {
                    MyCommond.ExportData<TxnTransaction>(ThisReceive, new ExportConfig()
                    {
                        Path1 = OtherDate_Path,
                        FileName = $"{Download_FileName}({Convert.ToDateTime(item).ToString("yyyyMMdd")})",
                        LastName = OtherDate_LastName.ToDescription(),
                    }, new TxnTransaction().ExportTitle_English(), itemTxn, Encoding.UTF8);
                }
                else
                {
                    MyCommond.ExportData(String.Format(OtherDateFullPath, Convert.ToDateTime(item).ToString("yyyyMMdd")), Encoding.UTF8, new TxnTransaction().ExportTitle_English());
                    foreach (var jtem in itemTxn)
                    {
                        MyCommond.ExportData(String.Format(OtherDateFullPath, Convert.ToDateTime(item).ToString("yyyyMMdd")), Encoding.UTF8, jtem.ToString());
                    }
                }
            }

            MyCommond.WriteLog(ThisReceive, $"結束");
            ExportloadData.FinishDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            UpLoaddata(CheckWrite_FullPath, LoadDataTitle, ExportloadData, CheckWrite_LastName);// 第二次寫入資料

            return true;
        }

        private bool UpLoaddata(string FullPath, string resutlTitle, LoadData result, FileType fileType)
        {
            try
            {
                List<LoadData> resultList = MyCommond.ToViewModel<LoadData>(ThisReceive, FullPath, resutlTitle, 1, true);
            StartCheck:
                List<LoadData> tempResult = resultList.ToList();
                var HaveWrite = tempResult.FindAll(x => x.FileName == result.FileName).ToList();

                if (HaveWrite.Count == 0)
                {
                    MyCommond.WriteLog(ThisReceive, $"未有相同資料");
                    goto ExportData;
                }
                else if (HaveWrite.Count == 1)
                {
                    bool TF1 = HaveWrite[0].FinishDate != "";
                    bool TF2 = HaveWrite[0].InstallDate == result.InstallDate;
                    bool TF3 = HaveWrite[0].StartDate == result.StartDate;
                    if (HaveWrite[0].FinishDate != "")
                    {
                        MyCommond.WriteLog(ThisReceive, $"已有相同資料");
                        return false;
                    }

                    if (TF2 && TF3)
                    {
                        var willDel = tempResult.Find(x => x.FileName == result.FileName && x.FinishDate == "");
                        tempResult.Remove(willDel);
                    }
                    else
                    {
                        MyCommond.WriteLog(ThisReceive, "發現異常載入");
                        return false;
                    }
                }
                else if (HaveWrite.Count > 1) 
                {
                    int tt = 0;
                    foreach (var ss in HaveWrite)
                    {
                        if (tt == 0)
                        {
                            MyCommond.WriteLog(ThisReceive, $"保留本筆資料:{ss.ToString()}");
                            continue;
                        }
                        MyCommond.WriteLog(ThisReceive, $"Warning! 刪除重複資料:{ss.ToString()}");
                        tempResult.Remove(ss);
                    }
                    goto StartCheck;
                }
                else 
                {

                }
                
            ExportData:
                tempResult.Add(result);
                ReWriteData<LoadData>(FullPath, resutlTitle, tempResult, fileType, Encoding.UTF8);
                return true;
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"Error! {ex.ToString()}");
                return false;
            }
            
        }

        private bool ReWriteData<TResult>(string FullPath, string Title, List<TResult> Data, FileType LastName, Encoding Ecode) where TResult : class, new()
        {
            int c = 0;

            if (LastName == FileType.CSV)
            {
                if (c == 0) { File.WriteAllText(FullPath, Title + Environment.NewLine, Ecode); c++; }
                if (Data == null) { return true; }
                foreach (var item in Data)
                {
                    File.AppendAllText(FullPath, item.ToString() + Environment.NewLine, Ecode);
                }
            }
            else if (LastName == FileType.Excel) { }
            else { }

            return true;
        }

        public void SetOd(bool _get, double _amt) { Od_get = _get; Od_amt = _amt;  }

        #endregion

        #region 編寫交易資料內容

        #region 原本

        private string Check_CardId(string cardId, string issuer, string samId)
        {
            string m_card = cardId;
            string m_str = $"CardId:{cardId,20}, Issuer:{issuer,4}, SamId:{samId,16}, ";
            if (issuer == "31")
            {

                m_card = $"{samId}{cardId.PadLeft(16, '0')}";
                m_str += $"NewCardId:{m_card,20}, ";
            }
            else
            {
                m_str += $"CardId:{cardId,20}, ";
            }

            if (globalCommond.Using_Mode == ExecutionMode.Debug) MyCommond.WriteLog(ThisReceive, m_str);
            return m_card;
        }

        private string Check_TxnType(string m_Data, ref string _result)
        {
            _result = m_Data;
            if (_result != "NULL")
            {
                List<String_String_String_String_Int> Result_List = globalCommond.Stock_Mrt_TxnType;
                if (Result_List.FindIndex(x => x.Num == Convert.ToInt32(m_Data)) > -1)
                {
                    _result = Result_List.Find(x => x.Num == Convert.ToInt32(m_Data)).ChName;
                }
            }
            return _result;
        }

        private string Check_TxnSubType(Ele_Iss Is_Id, string m_Data, ref string _result)
        {
            _result = m_Data;
            if (_result != "NULL")
            {
                List<String_String_String_String_Int> Result_List = new List<String_String_String_String_Int>();

                if (Is_Id == Ele_Iss.ECC) /*------------*/ { Result_List = globalCommond.Stock_Mrt_TxnTypeSub_Ecc; }
                else if (Is_Id == Ele_Iss.Ipass) /*-----*/ { Result_List = globalCommond.Stock_Mrt_TxnTypeSub_Ipass; }
                else if (Is_Id == Ele_Iss.Icash) /*-----*/ { Result_List = globalCommond.Stock_Mrt_TxnTypeSub_Icash; }
                else if (Is_Id == Ele_Iss.Mobile) /*----*/ { Result_List = globalCommond.Stock_Mrt_TxnTypeSub_Ecc; }
                else if (Is_Id == Ele_Iss.OwnTicket) /*-*/ { Result_List = globalCommond.Stock_Mrt_TxnTypeSub_Ecc; }
                else /*---------------------------------*/ { return _result; }

                if (Result_List.FindIndex(x => x.mCheck == "o" && x.Num == Convert.ToInt32(m_Data)) > -1)
                {
                    _result = Result_List.Find(x => x.mCheck == "o" && x.Num == Convert.ToInt32(m_Data)).ChName;
                }
            }
            return _result;
        }

        private UsuallyCommond.MyEnum.Ele_Iss Check_IssuerId(string m_Data, string TxnType, ref string _result, ref UsuallyCommond.MyEnum.Ele_Iss _result2)
        {
            _result = m_Data;
            if (_result != "NULL")
            {
                _result2 = UsuallyCommond.MyEnum.Ele_Iss.NA;

                if (Convert.ToInt16(TxnType) < 150 && Convert.ToInt16(_result) == 255) _result2 = UsuallyCommond.MyEnum.Ele_Iss.ECC;
                else if (Convert.ToInt16(_result) == 9) _result2 = UsuallyCommond.MyEnum.Ele_Iss.Ipass;
                else if (Convert.ToInt16(_result) == 11) _result2 = UsuallyCommond.MyEnum.Ele_Iss.Icash;
                else if (20 < Convert.ToInt16(_result) && Convert.ToInt16(_result) < 100) _result2 = UsuallyCommond.MyEnum.Ele_Iss.Mobile;
                else
                {
                    if (150 <= Convert.ToInt16(TxnType)) _result2 = UsuallyCommond.MyEnum.Ele_Iss.OwnTicket;
                    _result = _result2.ToDescription();
                    return _result2;
                }

                List<String_String_Int_Int_Int> Result_List = globalCommond.Stock_Mrt_IssureCode;
                if (-1 < Result_List.FindIndex(x => x.Num == Convert.ToInt32(m_Data))) _result = Result_List.Find(x => x.Num == Convert.ToInt32(m_Data)).ChName.Split('/')[0];
            }
            return _result2;
        }

        /**
        private string Check_IssuerId(string m_Data, string TxnType, ref string _result)
        {
            _result = m_Data;
            if (_result != "NULL")
            {
                if (Convert.ToInt16(TxnType) <= 150)
                {
                    List<String_String_String_Int> Result_List = globalCommond.Stock_Mrt_IssureCode;
                    if (Result_List.FindIndex(x => x.Num == Convert.ToInt32(m_Data)) > -1)
                    {
                        _result = Result_List.Find(x => x.Num == Convert.ToInt32(m_Data)).ChName.Split('/')[0];
                    }
                }
                else
                {
                    _result = "新北捷運";
                }
            }
            return _result;
        }
        */

        private string Check_Station(string m_Data, ref string _result)
        {
            _result = m_Data;
            if (_result != "NULL")
            {
                List<MrtStationList> Result_List = globalCommond.Stock_Mrt_Y_Station;
                if (Result_List.FindIndex(x => x.CodeName1 == m_Data) > -1)
                {
                    _result = Result_List.Find(x => x.CodeName1 == m_Data).StationName1;
                }
            }
            return _result;
        }

        private string Check_Fare(string m_Data, ref string _result)
        {
            _result = m_Data;
            if (_result != "NULL")
            {
                List<String_String_String_Int> Result_List = globalCommond.Stock_Mrt_FareProductType;
                if (Result_List.FindIndex(x => x.Num == Convert.ToInt32(m_Data)) > -1)
                {
                    _result = Result_List.Find(x => x.Num == Convert.ToInt32(m_Data)).ChName;
                }
            }
            return _result;
        }

        private string Check_Ticket(string m_Data, ref string _result)
        {
            _result = m_Data;
            if (_result != "NULL")
            {
                List<String_String_String_String_Int> Result_List = globalCommond.Stock_Mrt_TxnTypeSub_Ticket;
                if (Result_List.FindIndex(x => x.Num == Convert.ToInt32(m_Data)) > -1)
                {
                    _result = Result_List.Find(x => x.Num == Convert.ToInt32(m_Data)).ChName;
                }
            }
            return _result;
        }

        private string Check_IO(List<MrtStationOdList> results, E2 IO_, string m_Entry, string m_Exit)
        {
            try
            {
                var Find_Data = results.Find(x => x.Entry == m_Entry && x.Exit == m_Exit);
                     if (E2.Entry == IO_) { return Find_Data.Entry_Own; }
                else if (E2.Exit == IO_)  { return Find_Data.Exit_Own; }
                else                      { return "E1"; }
            }
            catch (Exception e)
            {
                     if (E2.Entry == IO_) {return m_Entry;}
                else if (E2.Exit == IO_)  {return m_Exit;}
                else                      {return "E2";}
            }
        }

        private string[] Check_IO_2(List<MrtStationOdList> results, string m_Entry, string m_Exit)
        {
            string[] vs = new string[2];
            try
            {
                var Find_Data = results.Find(x => x.Entry == m_Entry && x.Exit == m_Exit);
                vs[0] = Find_Data.Entry_Own;
            }
            catch (Exception e)
            {
                vs[0] = m_Entry;
            }
            try
            {
                var Find_Data = results.Find(x => x.Entry == m_Entry && x.Exit == m_Exit);
                vs[1] = Find_Data.Exit_Own;
            }
            catch (Exception e)
            {
                vs[1] = m_Exit;
            }

            return vs;
        }

        private string Check_TicketType(UsuallyCommond.MyEnum.Ele_Iss Issuer, string TxnTye, string SubTxn, string SubFare, string SubTicket, ref string ticket)
        {
            ticket = "";
            try
            {
                ticket = globalCommond.Stock_Mrt_FareProductType.Find(x => x.Num == Convert.ToInt16(SubFare)).VolumnCheck;
            }
            catch (Exception ex)
            {
                if (false) MyCommond.WriteLog(ThisReceive, $"輸入SubFare:{SubFare} - 未找到相符的悠遊卡定期票");
                try
                {
                    if (Issuer == UsuallyCommond.MyEnum.Ele_Iss.ECC)
                    {
                        ticket = globalCommond.Stock_Mrt_TxnTypeSub_Ecc.Find(x => x.Num == Convert.ToInt16(SubTxn)).VolumnCheck;
                    }
                    else if (Issuer == UsuallyCommond.MyEnum.Ele_Iss.Ipass)
                    {
                        ticket = globalCommond.Stock_Mrt_TxnTypeSub_Ipass.Find(x => x.Num == Convert.ToInt16(SubTxn)).VolumnCheck;
                    }
                    else if (Issuer == UsuallyCommond.MyEnum.Ele_Iss.Icash)
                    {
                        ticket = globalCommond.Stock_Mrt_TxnTypeSub_Icash.Find(x => x.Num == Convert.ToInt16(SubTxn)).VolumnCheck;
                    }
                    else if (Issuer == UsuallyCommond.MyEnum.Ele_Iss.Mobile)
                    {
                        ticket = globalCommond.Stock_Mrt_TxnTypeSub_Ecc.Find(x => x.Num == Convert.ToInt16(SubTxn)).VolumnCheck;
                    }
                    else if (Issuer == UsuallyCommond.MyEnum.Ele_Iss.OwnTicket)
                    {
                        ticket = globalCommond.Stock_Mrt_TxnTypeSub_Ticket.Find(x => x.Num == Convert.ToInt16(SubTicket)).VolumnCheck;
                    }
                    else
                    {
                        ticket = "";
                    }
                }
                catch (Exception ex2)
                {
                    if (false) MyCommond.WriteLog(ThisReceive, $"輸入Issuer:{Issuer.ToDescription()}, SubTxn:{SubTxn}, SubTicket:{SubTicket} - 未找到相符的子交易名稱");
                }
            }
            try
            {
                string s = ticket;
                ticket = globalCommond.Stock_Mrt_TicketType.Find(x => x.Code == s).TicketName;
            }
            catch (Exception ex)
            {
                ticket = "Error";
                MyCommond.WriteLog(ThisReceive, $"Error!{ex.Message}");
            }
            
            Console.WriteLine(ticket);
            return ticket;
        }

        /**
        private string Check_TicketType(string Issuer, string TxnTye, string SubTxn, string SubFare, string SubTicket, ref string ticket)
        {
            ticket = "";
            try
            {
                ticket = globalCommond.Stock_Mrt_FareProductType.Find(x => x.Num == Convert.ToInt16(SubFare)).VolumnCheck;
            }
            catch (Exception ex)
            {
                if (false) MyCommond.WriteLog(ThisReceive, $"輸入SubFare:{SubFare} - 未找到相符的悠遊卡定期票");
                try
                {

                    if (Convert.ToInt16(TxnTye) < 150)
                    {
                        if (Issuer == "255")
                        {
                            ticket = globalCommond.Stock_Mrt_TxnTypeSub_Ecc.Find(x => x.Num == Convert.ToInt16(SubTxn)).VolumnCheck;
                        }
                        else if (Issuer == "9")
                        {
                            ticket = globalCommond.Stock_Mrt_TxnTypeSub_Ipass.Find(x => x.Num == Convert.ToInt16(SubTxn)).VolumnCheck;
                        }
                        else if (Issuer == "11")
                        {
                            ticket = globalCommond.Stock_Mrt_TxnTypeSub_Icash.Find(x => x.Num == Convert.ToInt16(SubTxn)).VolumnCheck;
                        }
                    }
                    else
                    {
                        ticket = globalCommond.Stock_Mrt_TxnTypeSub_Ticket.Find(x => x.Num == Convert.ToInt16(SubTicket)).VolumnCheck;
                    }
                }
                catch (Exception ex2)
                {
                    if (false) MyCommond.WriteLog(ThisReceive, $"輸入Issuer:{Issuer}, SubTxn:{SubTxn}, SubTicket:{SubTicket} - 未找到相符的子交易名稱");
                }
            }
            try
            {
                string s = ticket;
                ticket = globalCommond.Stock_Mrt_TicketType.Find(x => x.Code == s).TicketName;
            }
            catch (Exception ex)
            {
                ticket = "Error";
                MyCommond.WriteLog(ThisReceive, $"Error!{ex.Message}");
            }

            Console.WriteLine(ticket);
            return ticket;
        }
*/

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
                var RouteList = globalCommond.Stock_Mrt_Y_OdList.ToList();
                var station = globalCommond.Stock_Mrt_Y_Station.ToList();

                var Route = RouteList.Find(x => x.Entry == m_Entry && x.Exit == m_Exit);

                string Byte1 = (((Route != null) ? ((Route.IsPass == "Y") ? "1" : "0") : "0"));
                string Byte2 = ((asdf.Contains(m_Exit)) ? "1" : "0");
                string EntryCompany = "0";
                string ExitCompany = "0";

                ErrorString += $"初始化變數\t";

                try
                {
                    ErrorString += $"判斷進站\t";

                    string tempI = station.Find(x => x.StationName1 == m_Entry).Company1;
                    if (tempI == "台北捷運") { EntryCompany = "1"; }
                    else if (tempI == "新北捷運") { EntryCompany = "2"; }
                }
                catch (Exception e)
                {
                    ErrorString += $"\n進站 - ERROR!\n{e.Message}\n";
                }

                try
                {
                    ErrorString += $"判斷出站\n";

                    string tempO = station.Find(x => x.StationName1 == m_Exit).Company1;
                    if (tempO == "台北捷運") { ExitCompany = "1"; }
                    else if (tempO == "新北捷運") { ExitCompany = "2"; }
                }
                catch (Exception e)
                {
                    ErrorString += $"\n出站 - ERROR!\n{e.Message}\n";
                }

                ErrorString += $"編寫代號 => Route.IsPass:{((Route != null) ? Route.IsPass : "E"),4}, EntryCompany:{EntryCompany,4}, ExitCompany:{ExitCompany,4}\n";
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
                TxnReportData_Count++;
                Console.WriteLine(ErrorString);
            }
        }

        #endregion 原本
        #region 新增

        private string AmountMark(string _error)
        {
            return $"{_error}";
        } // 營收標記

        private string ShoudPay(string _amt, string _entry, string _exit, ref double _num, ref string _error)
        {
            _error = "";
            _num = -1;
            try
            {
                var a = globalCommond.Stock_Mrt_Y_OdList.Find(x => x.Entry == _entry && x.Exit == _exit);
                if (a != null)
                {
                    _num = Convert.ToDouble(a.Amount_Full);
                }
            } 
            catch (Exception e)
            {
                MyCommond.WriteLog(ThisReceive, $@"Error!{e.Message}");
                _num = -1;
            }
            if (_num < 0) { _error += "A"; _num = Convert.ToDouble(_amt); }
            return $"{_num}";
        } // 應收金額(A)

        private string RealPay(string _amt, string _discStr, double _shoud, double _allpass, double _disc, double _off, ref double _num, ref string _error)
        {
            _num = -1;
            try
            {
                if (_discStr == "NULL")
                {
                    _num = ((_amt == "0") ? _shoud : Convert.ToDouble(_amt)) - _disc - _off;
                }
                else
                {
                    _num = ((_allpass > 0) ? 0 : Convert.ToDouble(_amt));
                }
            }
            catch (Exception e)
            {
                MyCommond.WriteLog(ThisReceive, $@"Error!{e.Message}");
                _num = -1;
            }
            if (_num < 0) { _error += "B"; _num = 0; }
            return $"{_num}";
        } // 實收金額(B)

        private string AllPassAmount(string _amt, string _ticket, ref double _num, ref string _error)
        {
            _num = -1;
            try
            {
                _num = ((_ticket == "公共運輸定期票") ? Convert.ToDouble(_amt) : 0);
            }
            catch (Exception e)
            {
                MyCommond.WriteLog(ThisReceive, $@"Error!{e.Message}");
                _num = -1;
            }
            if (_num < 0) { _error += "C"; _num = 0; }
            return $"{_num}";
        } // 公共運輸定期票票收(C)

        private string DiscAmount(double _shoud, string _disc, string _area, string _ticket, ref double _num, ref string _error)
        {
            _num = -1;
            List<string> _list_off_30 = new List<string> { "", "", "", "", "", "", "" };
            List<string> _list_off_40 = new List<string> { "兒童票", "", "", "", "", "", "" };
            List<string> _list_off_50 = new List<string> { "敬老卡", "愛心卡", "陪伴卡", "", "" };
            List<string> _list_off_50_ticket = new List<string> { "敬老票", "愛心票", "", "" };
            try
            {
                double _rr = 0;
                if (_disc != "NULL")
                {
                         if (_list_off_30.Contains(_ticket) || (_ticket == "兒童卡" && _area == "1"))
                    {
                        _rr = 0.3;
                    }
                    else if (_list_off_40.Contains(_ticket))
                    {
                        _rr = 0.4;
                    }
                    else if (_list_off_50.Contains(_ticket) || (_ticket == "兒童卡" && _area == "2") || (_ticket == "學生卡" && _area == "82"))
                    {
                        _rr = 0.5;
                    }
                    else
                    {

                    }
                }
                else
                {
                    if (_list_off_50_ticket.Contains(_ticket)) { _rr = 0.5; }
                }
                _num = Convert.ToDouble(_shoud) * _rr;
            }
            catch (Exception e)
            {
                MyCommond.WriteLog(ThisReceive, $@"Error!{e.Message}");
                _num = -1;
            }

            if (_num < 0 || (_disc != "NULL" && Convert.ToDouble(_disc) != (double)_num))
            {
                _error += "D";
                _num = 0;
            }

            return $"{_num}";
        } // 社福優惠(D)

        private string DiscPoin(string _point, string _ticket, double _dise, ref double _num, ref string _error)
        {
            _num = -1;
            try
            {
                if (_point != "NULL" && _dise > 0 && (_ticket == "敬老卡" || _ticket == "愛心卡"))
                { _num = Convert.ToDouble(_point); } else { _num = 0; }
            }
            catch (Exception e)
            {
                MyCommond.WriteLog(ThisReceive, $@"Error!{e.Message}");
                _num = -1;
            }
            if (_num < 0) { _error += "E"; _num = 0; }
            return $"{_num}";
        } // 社福點數(E)

        private string TrnNormal(string _trn, double _allpass, ref double _num, ref string _error)
        {
            _num = -1;
            try
            {
                if (_trn != "NULL")
                {
                    _num = ((_allpass == 0) ? Convert.ToDouble(_trn) : 0);
                }
            }
            catch (Exception e)
            {
                MyCommond.WriteLog(ThisReceive, $@"Error!{e.Message}");
                _num = -1;
            }
            if (_num < 0) { _error += "F"; _num = 0; }
            return $"{_num}";
        } // 轉乘優惠(F)

        private string TrnAllPass(string _trn, double _allpass, ref double _num, ref string _error)
        {
            _num = -1;
            try
            {
                if (_trn != "NULL")
                {
                    _num = ((_allpass != 0) ? Convert.ToDouble(_trn) : 0);
                }
            }
            catch (Exception e)
            {
                MyCommond.WriteLog(ThisReceive, $@"Error!{e.Message}");
                _num = -1;
            }
            if (_num < 0) { _error += "G"; _num = 0; }
            return $"{_num}";
        } // 定期票轉乘(G)

        private string Off_10(double _amt, double _disc, ref double _num, ref string _error)
        {
            _num = -1;
            if (_disc > 0)
            {
                try { _num = _amt * 0.1; }
                catch (Exception e)
                {
                    MyCommond.WriteLog(ThisReceive, $@"Error!{e.Message}");
                    _num = -1;
                }

            }
            else if (_disc == 0) { _num = 0; }
            if (_num < 0) { _error += "H"; _num = 0; }
            return $"{_num}";
        } // 社福吸收(H)

        private string DistanceRatio(string _entry, string _exit, ref double _num, ref string _error)
        {
            _num = -1;
            try
            {
                if (_entry != "NULL" && _exit != "NULL")
                {
                    var a = globalCommond.Stock_Mrt_Y_OdList.Find(x => x.Entry == _entry && x.Exit == _exit);
                    // Divide => 1 ~ 0.123456
                    if (a.IsPass == "Y" && a != null && a.MileAge_Scale != "" && a.MileAge_Scale == "1")
                    {
                        _num = 1;
                    }
                    else if (a.IsPass == "Y" && a != null && a.MileAge_Scale != "")
                    {
                        _num = Math.Round(Convert.ToDouble(a.MileAge_Scale), 4);
                        // _num = Convert.ToDouble(a.Substring(0, 6));
                    }
                    else if (a.IsPass == "N")
                    {
                        _num = 0;
                    }
                    else
                    {

                    }
                }
            }
            catch (Exception e)
            {
                MyCommond.WriteLog(ThisReceive, $@"Error!{e.Message}");
                _num = -1;
            }
            if (_num < 0) { _error += "I"; _num = 0; }
            return $"{_num}";
        } // 距離比例(I)

        private string Split_ShoudPay(double _amt, double _ratio, ref string _error)
        {
            double _num = -1;
            try
            {
                _num = _amt * _ratio;
            }
            catch (Exception e)
            {
                MyCommond.WriteLog(ThisReceive, $@"Error!{e.Message}");
                _num = -1;
            }
            if (_num < 0) { _error += "J"; _num = 0; }
            return $"{_num}";
        } // 應收金額(拆帳)(J)

        private string Split_RealPay(double _amt, double _ratio, ref string _error)
        {
            double _num = -1;
            try
            {
                _num = _amt * _ratio;
            }
            catch (Exception e)
            {
                MyCommond.WriteLog(ThisReceive, $@"Error!{e.Message}");
                _num = -1;
            }
            if (_num < 0) { _error += "K"; _num = 0; }
            return $"{_num}";
        } // 實收金額(拆帳)(K)

        private string Split_AllPassAmount(double _amt, double _ratio, ref string _error)
        {
            double _num = -1;
            try
            {
                _num = _amt * _ratio;
            }
            catch (Exception e)
            {
                MyCommond.WriteLog(ThisReceive, $@"Error!{e.Message}");
                _num = -1;
            }
            if (_num < 0) { _error += "L"; _num = 0; }
            return $"{_num}";
        } // 公共運輸定期票票收(拆帳)(L)

        private string Split_DiscAmount(double _amt, double _ratio, ref string _error)
        {
            double _num = -1;
            try
            {
                _num = _amt * _ratio;
            }
            catch (Exception e)
            {
                MyCommond.WriteLog(ThisReceive, $@"Error!{e.Message}");
                _num = -1;
            }
            if (_num < 0) { _error += "M"; _num = 0; }
            return $"{_num}";
        } // 社福優惠(拆帳)(M)

        private string Split_DiscPoint(double _amt, double _ratio, ref string _error)
        {
            double _num = -1;
            try
            {
                _num = _amt * _ratio;
            }
            catch (Exception e)
            {
                MyCommond.WriteLog(ThisReceive, $@"Error!{e.Message}");
                _num = -1;
            }
            if (_num < 0) { _error += "N"; _num = 0; }
            return $"{_num}";
        } // 社福點數(拆帳)(N)

        private string Split_TrnNormal(double _amt, double _ratio, ref string _error)
        {
            double _num = -1;
            try
            {
                _num = _amt * _ratio;
            }
            catch (Exception e)
            {
                MyCommond.WriteLog(ThisReceive, $@"Error!{e.Message}");
                _num = -1;
            }
            if (_num < 0) { _error += "O"; _num = 0; }
            return $"{_num}";
        } // 轉乘優惠(拆帳)(O)

        private string Split_TrnAllPass(double _amt, double _ratio, ref string _error)
        {
            double _num = -1;
            try
            {
                _num = _amt * _ratio;
            }
            catch (Exception e)
            {
                MyCommond.WriteLog(ThisReceive, $@"Error!{e.Message}");
                _num = -1;
            }
            if (_num < 0) { _error += "P"; _num = 0; }
            return $"{_num}";
        } // 定期票轉乘(拆帳)(P)

        private string Split_Off_10(double _amt, double _ratio, ref string _error)
        {
            double _num = -1;
            try
            {
                _num = _amt * _ratio;
            }
            catch (Exception e)
            {
                MyCommond.WriteLog(ThisReceive, $@"Error!{e.Message}");
                _num = -1;
            }
            if (_num < 0) { _error += "Q"; _num = 0; }
            return $"{_num}";
        } // 社福吸收(拆帳)(Q)

        private double ticketSidy(string _entry, string _exit, ref string _error)
        {
            try
            {
                var a = globalCommond.Stock_Mrt_Y_OdList.Find(x => x.Entry == _entry && x.Exit == _exit);
                if (a != null && a.Odamt != "") { return Convert.ToDouble(a.Odamt); }
            }
            catch (Exception ex)
            {
                _error += "R";
                MyCommond.WriteLog(ThisReceive, $@"ERROR!{ex.Message}");
            }
            return 0;
        } // 票差補貼(R)

        #endregion 新增

        #endregion

        #region 計算

        #region 已用

        #region 簡易資料(每日)

        private void Start_Easy_Analyze()
        {
            var me = MrtTxnClientSwitch.EzsyAnalyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #region 各交易資料

        private void Start_TxnType_Analyze()
        {
            var me = MrtTxnClientSwitch.TxnTypeAnalyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #region 進出站

        private void Start_IO_Analyze()
        {
            var me = MrtTxnClientSwitch.IOAnalyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #region 合併里程(延人公里)

        private void Start_Compare_Analyze_Old()
        {
            var me = MrtTxnClientSwitch.CompareAnalyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            for (int i = 0; i < CpuThread; i++)
            {
                MrtClient client = new MrtClient(this, me);
                client.Start();
                MrtClients.Add(client);
            }
        }

        private void Start_Compare_Analyze()
        {
            var me = MrtTxnClientSwitch.CompareAnalyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #region 轉乘

        private void Start_Trn_Analyze()
        {
            var me = MrtTxnClientSwitch.TrnAnalyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #region 通勤月票

        private void Start_TPass_Analyze()
        {
            var me = MrtTxnClientSwitch.TPassAnalyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #region 運量報表

        private void Start_VolumnReport_Analyze()
        {
            var me = MrtTxnClientSwitch.VolumnReportAnalyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #region 加值報表分析

        private void Start_AddValue_Analyze()
        {
            var me = MrtTxnClientSwitch.Day_AddValue;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #region 各家支付業者運量日報

        private void Start_EveryIssuer_Analyze()
        {
            var me = MrtTxnClientSwitch.Day_EveryIssuer;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #endregion

        #region 未用

        #region C10

        private void Start_C10_Analyze()
        {
            var me = MrtTxnClientSwitch.C10Analyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #region C11

        private void Start_C11_Analyze()
        {
            var me = MrtTxnClientSwitch.C11Analyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #region C12

        private void Start_C12_Analyze()
        {
            var me = MrtTxnClientSwitch.C12Analyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #region C13

        private void Start_C13_Analyze()
        {
            var me = MrtTxnClientSwitch.C13Analyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #region C14

        private void Start_C14_Analyze()
        {
            var me = MrtTxnClientSwitch.C14Analyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #region C15

        private void Start_C15_Analyze()
        {
            var me = MrtTxnClientSwitch.C15Analyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #region C16

        private void Start_C16_Analyze()
        {
            var me = MrtTxnClientSwitch.C16Analyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #region C17

        private void Start_C17_Analyze()
        {
            var me = MrtTxnClientSwitch.C17Analyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #region C18

        private void Start_C18_Analyze()
        {
            var me = MrtTxnClientSwitch.C18Analyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #region C19

        private void Start_C19_Analyze()
        {
            var me = MrtTxnClientSwitch.C19Analyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #region C20

        private void Start_C20_Analyze()
        {
            var me = MrtTxnClientSwitch.C20Analyze;
            MyCommond.WriteLog(ThisReceive, $"創建 {me.ToDescription()} 站台 * 1");
            MrtClient client = new MrtClient(this, me);
            client.Start();
            MrtClients.Add(client);
        }

        #endregion

        #endregion

        #endregion

        #region 輸出

        private void ExportTxnNormalFile()
        {
            while (true)
            {
                if (ResponseExportTxnNormal.Count() <= 0) continue;
                lock (ResponseExportTxnNormal)
                {
                    var item = ResponseExportTxnNormal[0];
                    string path = $@"{MyCommond.Path_Program}{MyCommond.Path_Tidy}{OperationStartDate:yyyy.MM}\TxnData{OperationStartDate:yyyyMMdd}.csv";
                    try
                    {
                        MyCommond.ExportData(path, Encoding.UTF8, item.ToString());
                    }
                    catch (Exception ex)
                    {

                    }
                    finally
                    {
                        ResponseExportTxnNormal.Remove(item);
                    }
                }
                Thread.Sleep(32);
            }
        }

        private void ExportTxnCompareFile()
        {
            while (true)
            {
                if (ResponseExportTxnCompare.Count() <= 0) continue;
                lock (ResponseExportTxnCompare)
                {
                    var item = ResponseExportTxnCompare[0];
                    string path = $@"{MyCommond.Path_Program}{MyCommond.Path_Tidy}{OperationStartDate:yyyy.MM}\TxnCompare{OperationStartDate:yyyyMMdd}.csv";
                    try
                    {
                        MyCommond.ExportData(path, Encoding.UTF8, item.ToString());
                    }
                    catch (Exception ex)
                    {

                    }
                    finally
                    {
                        ResponseExportTxnCompare.Remove(item);
                    }
                }
                Thread.Sleep(32);
            }
        }

        private void ExportTxnMegerFile()
        {
            while (true)
            {
                if (ResponseExportTxnMeger.Count() <= 0) continue;
                lock (ResponseExportTxnMeger)
                {
                    var item = ResponseExportTxnMeger[0];
                    string path = $@"{MyCommond.Path_Program}{MyCommond.Path_Tidy}{OperationStartDate:yyyy.MM}/TxnMeger{OperationStartDate:yyyyMMdd}.csv";
                    try
                    {
                        MyCommond.ExportData(path, Encoding.UTF8, item.ToString());
                    }
                    catch (Exception ex)
                    {

                    }
                    finally
                    {
                        ResponseExportTxnMeger.Remove(item);
                    }
                }
                Thread.Sleep(32);
            }
        }

        private void ExportTxnErrorFile()
        {
            while (true)
            {
                if (ResponseExportTxnError.Count() <= 0) continue;
                lock (ResponseExportTxnError)
                {
                    var item = ResponseExportTxnError[0];
                    string path = $@"{MyCommond.Path_Program}{MyCommond.Path_Tidy}{OperationStartDate:yyyy.MM}\TxnError{OperationStartDate:yyyyMMdd}.csv";
                    try
                    {
                        MyCommond.ExportData(path, Encoding.UTF8, item.ToString());
                    }
                    catch (Exception ex)
                    {

                    }
                    finally
                    {
                        ResponseExportTxnError.Remove(item);
                    }
                }
                Thread.Sleep(32);
            }
        }

        #endregion

    }

    public class ssss
    {
        public DateTime ddt { get; set; }
        public List<TxnTransaction> List_Txn_Transaction { get; set; }

    }
}
