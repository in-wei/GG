using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//--
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace QR_Code_FindOneDayTicket
{
    #region 參考
    public class QrViewList
    {
        public string OperationDate { get; set; }
        public string OperationLine { get; set; }
        public string Cooperate { get; set; }
        public int EntryCount { get; set; }
        public int ExitCount { get; set; }
        public int TicketCount { get; set; }

        public QrViewList()
        {
            EntryCount = 0;
            ExitCount = 0;
            TicketCount = 0;
        }

        public override string ToString()
        {
            return $"{OperationDate}" +
                $",{OperationLine}" +
                $",{Cooperate}" +
                $",{EntryCount}" +
                $",{ExitCount}" +
                $",{TicketCount}";
        }

        public string ExportTitle_English()
        {
            return $"OperationData" +
                $",OperationLine" +
                $",Cooperate" +
                $",EntryCount" +
                $",ExitCount" +
                $",TicketCount";
        }

        public string ExportTitle_zhTW()
        {
            return $"日期" +
                $",路線" +
                $",合作單位" +
                $",進站" +
                $",出站" +
                $",票券";
        }
    } // 回復資料參考

    public class TempData
    {
        public string No { get; set; }
        public string CardKind { get; set; }
        public string CardNumber_OutSide { get; set; }
        public string CardNumber_InSide { get; set; }
        public string CardNumber_Code { get; set; }
        public string ValidDate { get; set; }
        public string SupplyDate { get; set; }
        public string UsingDate { get; set; }
        public string Remark { get; set; }

        public TempData()
        {
            No = "";
            CardKind = "";
            CardNumber_OutSide = "";
            CardNumber_InSide = "";
            CardNumber_Code = "";
            ValidDate = "";
            SupplyDate = "";
            UsingDate = "";
            Remark = "";
        }

        public override string ToString()
        {
            return $"{No}" +
                $",{CardKind}" +
                $",{CardNumber_OutSide}" +
                $",{CardNumber_InSide}" +
                $",{CardNumber_Code}" +
                $",{ValidDate}" +
                $",{SupplyDate}" +
                $",{UsingDate}" +
                $",{Remark}";
        }

        public string ExportTitle_English()
        {
            return $"No" +
                $",CardKind" +
                $",CardNumber_OutSide" +
                $",CardNumber_InSide" +
                $",CardNumber_Code" +
                $",ValidDate" +
                $",SupplyDate" +
                $",UsingDate" +
                $",Remark";
        }

        public string ExportTitle_zhTW()
        {
            return $"項次" +
                $",票種" +
                $",外觀票號" +
                $",QR卡號" +
                $",編碼" +
                $",效期(天)" +
                $",提供日期" +
                $",開卡日期" +
                $",備註";
        }
    }   // 比對檔資料參考

    public class QRData
    {
        public string CardSN { get; set; }                  // 0
        public string CardKind { get; set; }                // 1
        public string OperationLine { get; set; }           // 2
        public string TxnType { get; set; }                 // 3
        public string TxnDateTime { get; set; }                 // 4
        public string EquipId { get; set; }                 // 5
        public string SerailNo { get; set; }                // 6
        public string ExpiryDate { get; set; }              // 7
        public string LocationID { get; set; }              // 8
        public string EmployeeID { get; set; }              // 9
        public string DataCode { get; set; }                // 10
        public string DataStatus { get; set; }              // 11
        public string IPASS_Identifier { get; set; }        // 12
        public string InstallDate { get; set; }             // 13
        public string ProcessingStatus { get; set; }        // 14
        public string ReplyInformation { get; set; }        // 15
        public string UpDateTime { get; set; }              // 16
        public string OutboundTimeForRefund { get; set; }   // 17
        public string AssignmentPlanDate { get; set; }      // 18
        public string SettlementAmount { get; set; }        // 19
        public string IPASS_TransactionNumber { get; set; } // 20
        public string IPASS_InstantFeedback { get; set; }   // 21
        public string SettlementDate { get; set; }          // 22
        public string ServiceProvider { get; set; }         // 23

        public QRData()
        {
            CardSN = "";
            CardKind = "";
            OperationLine = "";
            TxnType = "";
            TxnDateTime = "";
            EquipId = "";
            SerailNo = "";
            ExpiryDate = "";
            LocationID = "";
            EmployeeID = "";
            DataCode = "";
            DataStatus = "";
            IPASS_Identifier = "";
            InstallDate = "";
            ProcessingStatus = "";
            ReplyInformation = "";
            UpDateTime = "";
            OutboundTimeForRefund = "";
            AssignmentPlanDate = "";
            SettlementAmount = "";
            IPASS_TransactionNumber = "";
            IPASS_InstantFeedback = "";
            SettlementDate = "";
            ServiceProvider = "";
        }

        public override string ToString()
        {
            return $"{CardSN}" +                // 0
                $",{CardKind}" +                // 1
                $",{OperationLine}" +           // 2
                $",{TxnType}" +                 // 3
                $",{TxnDateTime}" +                 // 4
                $",{EquipId}" +                 // 5
                $",{SerailNo}" +                // 6
                $",{ExpiryDate}" +              // 7
                $",{LocationID}" +              // 8
                $",{EmployeeID}" +              // 9
                $",{DataCode}" +                // 10
                $",{DataStatus}" +              // 11
                $",{IPASS_Identifier}" +        // 12
                $",{InstallDate}" +             // 13
                $",{ProcessingStatus}" +        // 14
                $",{ReplyInformation}" +        // 15
                $",{UpDateTime}" +              // 16
                $",{OutboundTimeForRefund}" +   // 17
                $",{AssignmentPlanDate}" +      // 18
                $",{SettlementAmount}" +        // 19
                $",{IPASS_TransactionNumber}" + // 20
                $",{IPASS_InstantFeedback}" +   // 21
                $",{SettlementDate}" +          // 22
                $",{ServiceProvider}";          // 23
        }

        public string ExportTitle_English()
        {
            return $"CardSN" +                // 0
                $",CardKind" +                // 1
                $",OperationLine" +           // 2
                $",TxnType" +                 // 3
                $",TxnTime" +                 // 4
                $",EquipId" +                 // 5
                $",SerailNo" +                // 6
                $",ExpiryDate" +              // 7
                $",LocationID" +              // 8
                $",EmployeeID" +              // 9
                $",DataCode" +                // 10
                $",DataStatus" +              // 11
                $",IPASS_Identifier," +       // 12
                $",InstallDate" +             // 13
                $",ProcessingStatus" +        // 14
                $",ReplyInformation" +        // 15
                $",UpDateTime" +              // 16
                $",OutboundTimeForRefund" +   // 17
                $",AssignmentPlanDate" +      // 18
                $",SettlementAmount" +        // 19
                $",IPASS_TransactionNumber" + // 20
                $",IPASS_InstantFeedback" +   // 21
                $",SettlementDate" +          // 22
                $",ServiceProvider";          // 23
        }

        public string ExportTitle_zhTW()
        {
            return $"卡號" +         // 0
                $",電支業者" +       // 1
                $",營運路線" +       // 2
                $",交易類型" +       // 3
                $",設備交易時間" +   // 4
                $",設備編號" +       // 5
                $",交易序號" +       // 6
                $",到期日" +         // 7
                $",車站名稱" +       // 8
                $",員工編號" +       // 9
                $",資料" +           // 10
                $",資料狀態" +       // 11
                $",IPASS識別號" +    // 12
                $",存記錄時間" +     // 13
                $",處理狀態" +       // 14
                $",回覆資料" +       // 15
                $",更新時間" +       // 16
                $",退款的出站時間" + // 17
                $",分配計劃日期" +   // 18
                $",結算金額" +       // 19
                $",IPASS交易編號" +  // 20
                $",IPASS即時回饋" +  // 21
                $",結算日期" +       // 22
                $",服務提供者";      // 23
        }
    }     // 下載檔資料參考

    public class TempExcel
    {
        public string Cooperate { get; set; }
        public Excel.Application ExcelApp { get; set; }
        public Excel.Workbook ExcelWb { get; set; }
        public List<TempData> List_TempData { get; set; }
        public List<QRData> ConsistentDataList { get; set; }

        public TempExcel()
        {
            ExcelApp = new Excel.Application();
            //ExcelWb = new Excel.Workbook();
            List_TempData = new List<TempData>();
            ConsistentDataList = new List<QRData>();
        }

        public TempExcel(string G_Cooperate)
        {
            Cooperate = G_Cooperate;
            ExcelApp = new Excel.Application();
            //ExcelWb = new Excel.Workbook();
            List_TempData = new List<TempData>();
            ConsistentDataList = new List<QRData>();
        }
    }  // 比對檔資料合集參考
    #endregion

    public class QR_Code_FindOneDayTicket
    {
        #region 共用
        public string TxnDataFileName { get; set; }
        public DateTime OperationDate_Start { get; set; }
        public DateTime OperationDate { get; private set; }
        public DateTime OperationDate_End { get; set; }
        public string RunStep { get; set; }
        public bool IsRun { get; private set; }
        public string password { get; private set; }
        public List<QrViewList> qrViewList { get; private set; }     // 回復合併資料
        public List<TempExcel> tempExcels { get; private set; }      // 比對資料合集
        public List<QRData> TxnDataList { get; private set; }        // 下載資料合集
        #endregion

        #region 自用
        private MyCommonds MyCommond = new MyCommonds(true, false, true);
        private string ThisReceive;
        private bool TempExcelVisable = false;
        private string ExcelFileName_KKDAY = "Kkday_Temp.xlsx";
        private string ExcelFileName_NTPC = "NewTaiPeiCoin_Temp.xlsx";
        private string ExcelPassWord = "69278085";
        #endregion

        #region 合作廠商
        List<string> CooperateList = new List<string>();
        private string Cooperate_1 = "KKDAY";
        private string Cooperate_2 = "NTPC";
        #endregion

        public QR_Code_FindOneDayTicket()
        {
            ThisReceive = "QR_Code";
            string ProgramFullPath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName.ToString();
            string ProgramLastWriteTime = File.GetLastWriteTime(ProgramFullPath).ToString();
            MyCommond.WriteLog(ThisReceive, $"This Program Last Write Time： {ProgramLastWriteTime}");
            IsRun = false;
            TxnDataList = new List<QRData>();
            qrViewList = new List<QrViewList>();
            RunStep = "";
            password = "69278085";
        }

        private TempExcel AddTempStatus(string Cooperate) => new TempExcel() { Cooperate = Cooperate };
        
        public void Start()
        {
            IsRun = true;
            MyCommond.WriteLog(ThisReceive, $"準備一日票QR Code分析");
            qrViewList = new List<QrViewList>();

            System.Threading.Thread thread = new System.Threading.Thread(() =>
            {
                if ( true) LoadTxnData(); /*--------------------------------載入下載的資料*/
                if (TxnDataList.Count <= 0) { IsRun = false; return; } /*-----------------*/
                if ( true) tempDataStatus(); /*---------------------------產生合作廠商列表*/
                if ( true) LoadTempData(); /*---------------------------------載入比對資料*/
                if ( true) CompareData(); /*--------------------------------------比對資料*/
                if ( true) ExportCompareData(); /*----------------------------匯出比對資料*/
                if ( true) CloseListExcel(); /*-------------------------------關閉所有檔案*/

                System.Threading.Thread.Sleep(1000);
                RunStep = $"完成比對";
                IsRun = false;
            });
            thread.IsBackground = true;
            thread.Start();
        }

        private bool LoadTxnData(bool ExportLoadingLog = false)
        {
            MyCommond.WriteLog(ThisReceive, $"載入交易資料");
            RunStep = $"載入下載的資料";
            string path = @".\" + MyCommond.Path_Analyze;
            string FullPath = path + TxnDataFileName;

            string[] FullData;
            TxnDataList = new List<QRData>();

            int loopCount = 0;
            bool ReTryCheckFile = false;

        checkTxnDataFile:
            if (!MyCommond.CheckFile(FullPath))
            {
                MyCommond.WriteLog(ThisReceive, $"未發現檔案，重新導入路徑");
                FullPath = path + "Old/" + TxnDataFileName;
                ReTryCheckFile = true;
                if (loopCount++ < 1) { goto checkTxnDataFile; }
                else { goto EndLoadTxnData; }
            }
            else
            {
                MyCommond.WriteLog(ThisReceive, $"發現檔案，載入資料");
                loopCount = 0;
                FullData = File.ReadAllLines(FullPath);
            }

            if (FullData.Length <= 0) return false;

            MyCommond.WriteLog(ThisReceive, $"確定有資料載入，開始寫入記憶體");
            string TxnTitle = "";
            string[] ListTitle = new QRData().ExportTitle_zhTW().Split(',');  // 取得指定欄位

            int RunCount = 0;
            foreach (var Data in FullData)
            {
                if (ExportLoadingLog) { MyCommond.WriteLog(ThisReceive, $"ReadLine:{Data.ToString()}"); } // 讀取資料行
                if (TxnTitle == "") { TxnTitle = Data.ToString(); goto NextData; }  // 寫入欄位資訊

                string[] mData = new string[ListTitle.Length];
                int DataPut = 0;

                try
                {
                    string[] split_1 = { ",{" };
                    string[] data_1 = Data.Split(split_1, System.StringSplitOptions.RemoveEmptyEntries);  // 0 - 前半資料  1 - 回饋資料+後半資料

                    string[] split_2 = { "}," };
                    string[] data_2 = data_1[1].Split(split_2, System.StringSplitOptions.RemoveEmptyEntries);   // 0 - 回饋資料   1 - 後半資料

                    string[] data_1_1 = data_1[0].Split(',');
                    string[] data_2_1 = data_2[1].Split(',');

                    foreach (var itemput in data_1_1) { mData[DataPut++] = itemput; }
                    mData[DataPut++] = @"{" + data_2[0].ToString() + @"}";
                    foreach (var itemput in data_2_1) { mData[DataPut++] = itemput; }
                }
                catch (Exception ex)
                {
                    mData = Data.Split(',');
                }
                finally
                {

                } // 資料處裡

                QRData tr = MyCommond.T2<QRData>(ThisReceive, new QRData().ExportTitle_zhTW(), TxnTitle.ToString(), mData); // 資料寫入記憶體
                TxnDataList.Add(tr);
            NextData:
                RunCount++;
            }

            string path1 = $"{MyCommond.Path_Report}/一日票QR Code/{Convert.ToDateTime(OperationDate_Start).ToString("yyyyMMdd")}/";
            string filename = MyCommond.CheckAndReName(path1, $"{Convert.ToDateTime(OperationDate_Start).ToString("yyyyMMdd")}-{Convert.ToDateTime(OperationDate_End).ToString("yyyyMMdd")}_行動支付即時交易明細_QrCode_全線", "xlsx");

            if (OperationDate_Start.AddDays(1) != OperationDate_End) { goto EndLoadTxnData; } // 比對大於一天就不輸出明細
            MyCommond.ExportData2(ThisReceive, new ExportConfig()
            {
                Path1 = path1,
                DTime = new DateTime(),
                FirstName = "",
                FileName = filename.Split('.')[0],
                SubName = "",
                OpLine = "",
                LastName = "xlsx",
            }, new QRData().ExportTitle_zhTW(), TxnDataList, Encoding.UTF8);

            if (!ReTryCheckFile) { MyCommond.MoveFile(path, path + "Old/", TxnDataFileName); }  // 將下載的檔案一到其他目錄

        EndLoadTxnData:
            MyCommond.WriteLog(ThisReceive, $"載入完成");
            return true;
        } // 載入交易資料
        
        private void tempDataStatus()
        {
            RunStep = $"產生合作廠商列表";
            tempExcels = new List<TempExcel>();
            //tempExcels.Add(AddTempStatus(Cooperate_1));
            //tempExcels.Add(AddTempStatus(Cooperate_2));
            CooperateList = new List<string>() { Cooperate_1, Cooperate_2 };
            foreach(var item in CooperateList)
            {
                //tempExcels.Add(AddTempStatus(item));
                tempExcels.Add(new TempExcel(item));
            }
        }

        private bool LoadTempData()
        {
            MyCommond.WriteLog(ThisReceive, $"載入各合作廠商所有票券");
            RunStep = $"載入比對資料";
            foreach (var item in CooperateList)
            {
                var tempListData = tempExcels.Find(x => x.Cooperate == item);
                string path = MyCommond.Path_Program + MyCommond.Path_Template;
                string FullPath = path;
                if (item == Cooperate_1) { FullPath += ExcelFileName_KKDAY; }
                else if (item == Cooperate_2) { FullPath += ExcelFileName_NTPC; }
                else { }
                bool checkFile = MyCommond.CheckFile(FullPath);
                //string password = "69278085";
                MyCommond.WriteLog(ThisReceive, $"合作廠商名稱  :{item}");
                MyCommond.WriteLog(ThisReceive, $"比對檔完整路徑:{FullPath}");
                MyCommond.WriteLog(ThisReceive, $"是否有此檔案  :{ ((checkFile) ? "有" : "無")}");
                if (!checkFile) return false;
                try
                {
                    tempListData.ExcelWb = tempListData.ExcelApp.Workbooks.Open(Filename: FullPath, ReadOnly: false, Password: password);
                    tempListData.ExcelApp.Visible = TempExcelVisable;
                    Excel.Worksheet xlWs = (Excel.Worksheet)tempListData.ExcelWb.Worksheets[1];
                    Excel.Range xlWr = xlWs.UsedRange;
                    string titleGather = "";
                    int ExitLoop = 0;
                    MyCommond.WriteLog(ThisReceive, $"取得檔案欄位名稱");
                    for (int i = 1; i <= 25; i++)
                    {
                        var CellValue = Convert.ToString((xlWr[1, i] as Excel.Range).Value2);

                        if (CellValue != null)
                        {
                            titleGather += CellValue + ",";
                        }
                        else
                        {
                            if (ExitLoop > 10) break;
                            ExitLoop++;
                        }
                    }
                    titleGather = titleGather.Substring(0, titleGather.Length - 1);
                    int titleCount = titleGather.Split(',').Length;
                    ExitLoop = 0;
                    MyCommond.WriteLog(ThisReceive, $"將比對檔資料寫入記憶體");
                    for (int ir = 2; ir < 10000; ir++)
                    {
                        string[] tempRowData = new string[titleCount];
                        for (int ic = 1; ic <= titleCount; ic++)
                        {
                            var CellValue = Convert.ToString((xlWr[ir, ic] as Excel.Range).Text);
                            if (CellValue != null && CellValue != "") { tempRowData[ic - 1] = CellValue; }
                        }

                        bool checkArrage = false;
                        foreach (var jtem in tempRowData) 
                        {
                            if (jtem == null) continue;
                            checkArrage = true;
                            break;
                        }

                        if (checkArrage)
                        {
                            TempData td = MyCommond.T2<TempData>(ThisReceive, new TempData().ExportTitle_zhTW(), titleGather, tempRowData);
                            tempListData.List_TempData.Add(td);
                        }
                        else
                        {
                            if (ExitLoop > 10) { break; }
                            ExitLoop++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"Error!\n{ex.Message}");
                }
                MyCommond.WriteLog(ThisReceive, $"讀取 {item} 比對檔資料完成");
            }
            return true;
        } // 載入比對資料

        private bool CompareData()
        {
            RunStep = $"比對資料";
            foreach (var item in CooperateList)
            {
                var s1 = tempExcels.Find(x => x.Cooperate == item);
                
                //取出比對檔的卡號
                var s2 = (from p in s1.List_TempData select p.CardNumber_InSide).ToList();

                //找出相符卡號的交易資料
                var s4 = (from p in TxnDataList where s2.Contains(p.CardSN) && s1.List_TempData.Find(x => x.CardNumber_InSide == p.CardSN).CardNumber_Code == p.DataCode select p).ToList();
                
                //var s4 = (from p in TxnDataList where s2.Contains(p.CardSN) select p).ToList();

                foreach (var jtem in s4)
                {
                    tempExcels.Find(x => x.Cooperate == item).List_TempData.Find(x => x.CardNumber_InSide == jtem.CardSN).UsingDate = Convert.ToDateTime(jtem.TxnDateTime).ToString("yyyy/MM/dd");
                }
                s1.ConsistentDataList = s4;
            }
            return false;
        } // 比對資料

        private bool ExportCompareData()
        {
            // 依照不同通路去切日期分析資料

            RunStep = $"匯出比對資料";
            string ExportPath = $"{MyCommond.Path_Report}/一日票QR Code/";
            foreach (var item in CooperateList)
            {
                var CompareCooperate = tempExcels.Find(x => x.Cooperate == item);

                for (DateTime OPDate = OperationDate_Start; OPDate < OperationDate_End; OPDate = OPDate.AddDays(1))
                {
                    var OnDate_Data = CompareCooperate.ConsistentDataList.FindAll(x => Convert.ToDateTime(x.TxnDateTime) >= OPDate && Convert.ToDateTime(x.TxnDateTime) < OPDate.AddDays(1)).ToList();

                    #region 輸出文件

                    if (item == "NTPC")
                    {
                        string path = $"{ExportPath}/{OPDate:yyyyMMdd}/";
                        string filename = $"{OPDate.ToString("yyyyMMdd")}";
                        string ExportFileFullPath = path + MyCommond.CheckAndReName(path, filename, "csv");
                        MyCommond.ExportData(ExportFileFullPath, Encoding.UTF8, "卡號,設備交易時間,資料");

                        var NtpcCardList = (from p in OnDate_Data group p by new {p.CardSN, p.DataCode} into g select new string[] { g.Key.CardSN, g.Key.DataCode}).ToList();
                        
                        foreach (var jtem in NtpcCardList)
                        {
                            var exportData = CompareCooperate.ConsistentDataList.Find(x => x.CardSN == jtem[0] && x.DataCode == jtem[1]);

                            string sss = $"{jtem[0]},{exportData.TxnDateTime},{jtem[1]}";

                            MyCommond.ExportData(ExportFileFullPath, Encoding.UTF8, sss);
                        }
                    }  // 輸出回傳給新北政府 以新北幣兌換一日票 且有來系統使用的卡號

                    if (OnDate_Data.Count > 0)
                    {
                        MyCommond.ExportData2<QRData>(ThisReceive, new ExportConfig()
                        {
                            Path1 = $"{ExportPath}/{OPDate:yyyyMMdd}/",
                            DTime = OPDate,
                            FirstName = "",
                            FileName = "比對明細",
                            SubName = "",
                            OpLine = item,
                            LastName = "xlsx",
                        }, new QRData().ExportTitle_zhTW(), OnDate_Data, Encoding.UTF8); // 輸出比對到符合的資料
                    }

                    #endregion 輸出文件

                    if (true)
                    {
                        foreach (var jtem in OnDate_Data)
                        {
                            bool HaveData = qrViewList.FindAll(x => 
                                x.OperationDate == OPDate.ToString("yyyyMMdd") && 
                                x.OperationLine == jtem.OperationLine && 
                                x.Cooperate == item).Count > 0;

                            if (qrViewList == null || !HaveData)
                            {
                                qrViewList.Add(new QrViewList()
                                {
                                    OperationDate = OPDate.ToString("yyyyMMdd"),
                                    OperationLine = jtem.OperationLine,
                                    Cooperate = item,
                                    EntryCount = ((jtem.TxnType == "進站") ? 1 : 0),
                                    ExitCount = ((jtem.TxnType == "出站") ? 1 : 0),
                                    TicketCount = 0,
                                });
                            }
                            else
                            {
                                var VL_1 = qrViewList.Find(x => x.OperationDate == OPDate.ToString("yyyyMMdd") && x.OperationLine == jtem.OperationLine && x.Cooperate == item);

                                VL_1.EntryCount = ((jtem.TxnType == "進站") ? VL_1.EntryCount + 1 : VL_1.EntryCount);
                                VL_1.ExitCount  = ((jtem.TxnType == "出站") ? VL_1.ExitCount + 1  : VL_1.ExitCount );

                            }
                        }

                        var OnDate_Data_CardSN = OnDate_Data.GroupBy(x => x.CardSN).ToList();
                        foreach (var ktem in OnDate_Data_CardSN)
                        {
                            var Txn_Line = OnDate_Data.FindLast(x => x.CardSN == ktem.Key);
                            var VL_2 = qrViewList.Find(x => x.OperationDate == OPDate.ToString("yyyyMMdd") && x.OperationLine == Txn_Line.OperationLine && x.Cooperate == item);

                            VL_2.TicketCount = VL_2.TicketCount + 1;
                        }

                    }
                    else
                    {
                        var CompareLine = (from p in CompareCooperate.ConsistentDataList group p by new { p.OperationLine } into g select g.Key.OperationLine).ToList();

                        foreach (var jtem in CompareLine)
                        {
                            var GetTxnData = (from p in CompareCooperate.ConsistentDataList where p.OperationLine == jtem && Convert.ToDateTime(p.TxnDateTime) >= OPDate && Convert.ToDateTime(p.TxnDateTime) < OPDate.AddDays(1) select p).ToList();

                            var TV_Entry = (from p in GetTxnData where p.TxnType == "進站" select p).ToList();
                            var TV_Exit = (from p in GetTxnData where p.TxnType == "出站" select p).ToList();
                            var TV_Ticket = (from p in GetTxnData group p.CardSN by new { p.CardSN } into g select g.Key.CardSN).ToList();

                            if (TV_Entry.Count == 0 && TV_Exit.Count == 0 && TV_Ticket.Count == 0) { continue; }

                            QrViewList qrViewList1 = new QrViewList()
                            {
                                OperationDate = OPDate.ToString("yyyyMMdd"),
                                OperationLine = jtem,
                                Cooperate = item,
                                EntryCount = TV_Entry.Count,
                                ExitCount = TV_Exit.Count,
                                TicketCount = TV_Ticket.Count,
                            };

                            qrViewList.Add(qrViewList1);
                        }  // 將所有比對到的資料寫成一個List給UI抓來輸出到畫面

                    }

                }
            }
            if (OperationDate_Start.AddDays(1) != OperationDate_End) { return false; } // 比對大於一天就不輸出明細
            MyCommond.ExportData2<QrViewList>(ThisReceive, new ExportConfig()
            {
                Path1 = $"{ExportPath}/{OperationDate_Start:yyyyMMdd}/",
                DTime = OperationDate_Start,
                FirstName = $"",
                FileName = "簡易資料輸出",
                SubName = "",
                OpLine = "",
                LastName = "xlsx",
            }, new QrViewList().ExportTitle_zhTW(), qrViewList, Encoding.UTF8); // 輸出比對到符合的資料

            return false;
        } // 匯出比對資料

        private bool CloseListExcel()
        {
            MyCommond.WriteLog(ThisReceive, $"準備儲存並關閉檔案");
            RunStep = $"關閉所有檔案";
            foreach (var item in tempExcels)
            {
                try
                {
                    ReWriteExcel(item.Cooperate);
                    if (true) item.ExcelWb.Save();
                    item.ExcelWb.Close();
                    item.ExcelApp.Quit();
                    Marshal.ReleaseComObject(item.ExcelApp);
                    Marshal.ReleaseComObject(item.ExcelWb);
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
                }
            }
            GC.Collect();
            return false;
        } // 關閉所有檔案

        private int FindStringArrage(string text, string[] data)
        {
            int re = 1;
            foreach (var item in data)
            {
                if (item == text) { break; }
                else { re++; }
            }
            return re;
        } // 搜尋欄位位置

        private bool ReWriteExcel(string Cooperate)
        {
            try
            {
                MyCommond.WriteLog(ThisReceive, $"");
                var tempListData = tempExcels.Find(x => x.Cooperate == Cooperate);
                var UsingCardList = 
                    (
                        from p in tempListData.List_TempData 
                        where
                               (/*p.UsingDate != null || */p.UsingDate != "") 
                            && Convert.ToDateTime(Convert.ToDateTime(p.UsingDate).ToString("yyyy/MM/dd")) >= Convert.ToDateTime(Convert.ToDateTime(OperationDate_Start).ToString("yyyy/MM/dd")) 
                            && Convert.ToDateTime(Convert.ToDateTime(p.UsingDate).ToString("yyyy/MM/dd")) <  Convert.ToDateTime(Convert.ToDateTime(OperationDate_End).ToString("yyyy/MM/dd")) 
                        select p
                    ).ToList();
                //var tempD1 = tempListData.ConsistentDataList.FindAll(x => Convert.ToDateTime(x.TxnDateTime).ToString("yyyy/MM/dd") == OperationDate.ToString("yyyy/MM/dd"));

                var title = new TempData().ExportTitle_zhTW().Split(',');

                foreach (var item in UsingCardList)
                {
                    try
                    {
                        var tempDataIndex = tempListData.List_TempData.FindIndex(x => x.CardNumber_InSide == item.CardNumber_InSide && x.CardNumber_Code == item.CardNumber_Code) + 2;
                        if (tempDataIndex > -1)
                        {
                            var xlWs = tempListData.ExcelWb.Worksheets[1];
                            var xlWr = xlWs.UsedRange;

                            if (!true) { /*xlWr.Cell(tempDataIndex, FindStringArrage("項次"     , title)).Style.NumberFormat.Format = "@";*/ xlWr.Cells[tempDataIndex, FindStringArrage("項次"     , title)] = item.No; }
                            if (!true) { /*xlWr.Cell(tempDataIndex, FindStringArrage("票種"     , title)).Style.NumberFormat.Format = "@";*/ xlWr.Cells[tempDataIndex, FindStringArrage("票種"     , title)] = item.CardKind; }
                            if (!true) { /*xlWr.Cell(tempDataIndex, FindStringArrage("外觀票號" , title)).Style.NumberFormat.Format = "@";*/ xlWr.Cells[tempDataIndex, FindStringArrage("外觀票號" , title)] = item.CardNumber_OutSide; }
                            if (!true) { /*xlWr.Cell(tempDataIndex, FindStringArrage("QR卡號"   , title)).Style.NumberFormat.Format = "@";*/ xlWr.Cells[tempDataIndex, FindStringArrage("QR卡號"   , title)] = item.CardNumber_InSide; }
                            if (!true) { /*xlWr.Cell(tempDataIndex, FindStringArrage("編號"     , title)).Style.NumberFormat.Format = "@";*/ xlWr.Cells[tempDataIndex, FindStringArrage("編號"     , title)] = item.CardNumber_Code; }
                            if (!true) { /*xlWr.Cell(tempDataIndex, FindStringArrage("效期(天)" , title)).Style.NumberFormat.Format = "@";*/ xlWr.Cells[tempDataIndex, FindStringArrage("效期(天)" , title)] = item.ValidDate; }
                            if (!true) { /*xlWr.Cell(tempDataIndex, FindStringArrage("提供日期" , title)).Style.NumberFormat.Format = "@";*/ xlWr.Cells[tempDataIndex, FindStringArrage("提供日期" , title)] = item.SupplyDate; }
                            if ( true) { /*xlWr.Cell(tempDataIndex, FindStringArrage("開卡日期" , title)).Style.NumberFormat.Format = "@";*/ xlWr.Cells[tempDataIndex, FindStringArrage("開卡日期" , title)] = Convert.ToDateTime(item.UsingDate).ToString("yyyy/MM/dd"); }
                            if (!true) { /*xlWr.Cell(tempDataIndex, FindStringArrage("備註"     , title)).Style.NumberFormat.Format = "@";*/ xlWr.Cells[tempDataIndex, FindStringArrage("備註"     , title)] = item.Remark; }
                        }
                    }
                    catch (Exception ex)
                    {
                        MyCommond.WriteLog(ThisReceive, $"ERROR!Try(1)\n{ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!Try(0)\n{ex.Message}");
            }
            return false;
        } // 寫入比對完的資料

    }
}
