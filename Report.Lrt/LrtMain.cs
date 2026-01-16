using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//-
using System.IO;
using System.Threading;
using UsuallyCommond.MyEnum;
//using Report.Lrt.ViewModel;
//using Report.Lrt.Load;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using GlobalCommond;
using GlobalCommond.ViewModel;

namespace Report.Lrt
{
    public class LrtMain
    {
        #region 變數

        #region 設定(公開)
        public ExecutionMode Using_Mode { get; set; }
        public bool IsRunning { get; private set; }
        public List<LrtClient> LrtClients { get; set; }
        public DateTime OperationStartDate { get; set; }
        public DateTime OperationEndDate { get; set; }
        public bool Thread_Use { get; set; }

        public bool VA_Lock = false;

        public List<PaoTicketSaleList> paoTicketSaleLists { get; set; } //202301_Pao.csv
        public List<MobilePay> ReadMobilePayData { get; set; }  //MobileData_202301.csv

        public string ClientLock { get; set; }
        #endregion

        #region 檔案

        public string Config_Subidy = "Subsidy.csv";
        public string Config_Report = "Report.csv";
        public string Config_StationList = "StationList.csv";

        #region 組態檔
        public readonly string Config_Setting = @"setting.config";   //存放程式設定檔
        public readonly string Config_OperationLine = @"Line.csv"; //存放營運中的路線
        public readonly string Config_PaoStation = @"StationAmtList.txt";    //存放班結表版本的檔案
        public readonly string Config_PaoList = @"Pao.csv";      //PAO車站設定
        public readonly string Config_StationFile = @"{0}/Station.csv";  //存放詳細線路名稱的檔案
        public readonly string Config_ReportFile = @"{0}/Report.csv";    //產出檔設定
        public readonly string Config_Subsidy = @"{0}/Subsidy.csv";  //票差補貼設定
        public readonly string Config_SqlList = @"SqlList.csv";      //SQL語法設定
        public readonly string Config_MobileCard = "MobileCard.csv";
        #endregion 

        #region 自用資料庫檔名
        public readonly string FileName_GroupTicket = "GroupTicket";
        public readonly string FileName_PaoCsv = "PaoData";
        #endregion

        #region 匯入資料檔名
        public readonly string FileName_QrCode = "QRCode每日交易表-{0}.csv";                             //QRCode每日交易表-2022-10-15.csv
        public readonly string FileName_CreditCard = "信用卡每日交易表-{0}.csv";                         //信用卡每日交易表-2022-11-06.csv
        public readonly string FileName_PaoStation = "{0}輕軌候車站營收班結記錄表{1}-{2}.xls";          //淡海輕軌候車站營收班結記錄表V26-8031.xls
        public readonly string FileName_PaoTicketControl = "票卡存管統計表-{0)月-{1}.xlsx";              //票卡存管統計表-111.08月-V01
        #endregion

        #region 匯出資料檔名

        #region 日
        /*特殊 需同檔案一同修改*/
        public readonly string FileName_DayRideAmt = @"{0}_營運資料統計表(總運量、總營收)_{1}.xlsx";         //1120101_營運資料統計表(總運量、總營收).xlsx
        public readonly string FileName_DayStationTime = @"{0}_每日全線各站分時運量日報表_{1}.xlsx";         //1120101_每日全線各站分時運量日報表.xlsx
        public readonly string FileName_DayTrafficVolume = @"{0}_票卡進出站運量日報表_{1}.xlsx";             //1120101_票卡進出站運量日報表.xlsx
        public readonly string FileName_DayRideAllList = @"{0}_日運量總表_{1}.xlsx";                         //1120101_日運量總表.xlsx(停用)
        public readonly string FileName_DayStation = @"{0}_每日起訖總表_{1}.xlsx";                           //1120101_每日起訖總表.xlsx
        public readonly string FileName_DayAmt = @"{0}_營收日報表_{1}.xlsx";                                 //1120101_營收日報表.xlsx
        public readonly string FileName_DayOd = @"{0}_每日起訖總表_{1}.xlsx";                                //1120101_每日起訖總表.xlsx
        public readonly string FileName_DayEquipAmt = @"{0}_設備營收日報表_{1}.xlsx";                        //1120101_設備營收日報表.xlsx
        #endregion

        #region 月
        public readonly string FileName_MonthTrafficVolume = @"{0}_月運量統計表_{1}.xlsx";                   //1120101_月運量統計表.xlsx(停用)
        public readonly string FileName_MonthOD = @"{0}_每月起訖總表_{1}.xlsx";                              //1120101_每月起訖總表.xlsx
        public readonly string FileName_MonthTicketStation = @"{0}_各家票卡運量月報表(車站)_{1}.xlsx";       //1120101_各家票卡運量月報表(車站).xlsx
        public readonly string FileName_MonthTicketDay = @"{0}_各家票卡運量月報表(天期)_{1}.xlsx";           //1120101_各家票卡運量月報表(天期).xlsx
        public readonly string FileName_MonthLrtTicketStation = @"{0}_自有票卡運量月報表(車站)_{1}.xlsx";    //1120101_自有票卡運量月報表(車站).xlsx
        public readonly string FileName_MonthLrtTicketDay = @"{0}_自有票卡運量月報表(天期)_{1}.xlsx";        //1120101_自有票卡運量月報表(天期).xlsx
        public readonly string FileName_MonthTickerStationAmt = @"{0}_營收月報表 (車站)_{1}.xlsx";           //1120101_營收月報表(車站).xlsx
        public readonly string FileName_MonthTicketDayAmt = @"{0}_營收月報表(天期)_{1}.xlsx";                //1120101_營收月報表(天期).xlsx
        public readonly string FileName_MonthEquipStationAmt = @"{0}_設備營收月報表(車站)_{1}.xlsx";         //1120101_設備營收月報表(車站).xlsx
        public readonly string FileName_MonthEquipDayAmt = @"{0}_設備營收月報表(天期)_{1}.xlsx";             //1120101_設備營收月報表(天期).xlsx
        #endregion

        #endregion

        #endregion

        #region 設定(封閉)

        private int CpuThread;
        public MyCommonds MyCommond;
        private string ThisReceive = "MainLrt";
        public GlobalCommonds globalCommond { get; set; }
        Thread threadLoading;

        bool CheckPaoCsv;
        #region 路徑

        public readonly string Path_PaoData = $@"PaoData\";
        public readonly string StationAmtConfig = @"StationAmtList.txt";

        #endregion

        #endregion

        #endregion

        public LrtMain()
        {
            MyCommond = new MyCommonds();
            string ProgramFullPath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName.ToString();
            string ProgramLastWriteTime = File.GetLastWriteTime(ProgramFullPath).ToString();
            MyCommond.WriteLog(ThisReceive, $"This Program Last Write Time： {ProgramLastWriteTime}");
            IsRunning = false;
            ClientLock = "";
        }

        public void SetMyCommand()
        {
            globalCommond.Loading_Normal_Finish = false;
            MyCommond = globalCommond.MyCommond;
        }

        #region 功能

        private void GetCpuThread()
        {
            CpuThread = Environment.ProcessorCount - 1;
            Console.WriteLine($"Processors {CpuThread}");
        }

        public void Start(List<CheckBoxs> CB, List<PaoTicketSaleList> FormData)
        {
            LrtClients = new List<LrtClient>();
            MyCommond.CheckFolder(MyCommond.Path_Program + Path_PaoData + @"Old\");
            int searchStatus = (int)CB.Find(x => x.Status == true).Methode;
            string PaoDataPath = MyCommond.Path_Program + Path_PaoData + $"{Convert.ToString(OperationStartDate.ToString("yyyyMM"))}_{FileName_PaoCsv}_{globalCommond.OperationCode_String}.csv";
            paoTicketSaleLists = MyCommond.ToViewModel<PaoTicketSaleList>(ThisReceive, PaoDataPath, (new PaoTicketSaleList()).ExportTitle_zhTW(), 1, true);
            if (searchStatus > 8)
            {
                string MobileDataPath_0 = $@"{MyCommond.Path_Program}{MyCommond.Path_Analyze}MobileData_{OperationStartDate:yyyy-MM}.csv";
                string MobileDataPath = MobileDataPath_0;
                if (!MyCommond.CheckSimilarFile(ref MobileDataPath))
                {
                    var CheckMessageBox = MessageBox.Show($"發現電支檔案已經合併過，請確認是否要覆蓋該檔案", "注意", MessageBoxButtons.YesNoCancel);
                    MyCommond.WriteLog(ThisReceive, $"取得使用者回覆：{CheckMessageBox.ToString()}");
                    if (CheckMessageBox == DialogResult.Yes) { MobileData2File(); /*重新合併1個月的行動支付檔案*/ }
                    else if (CheckMessageBox == DialogResult.No) { /*以檢測到的檔案執行*/ }
                    else { return; /*結束執行*/ }
                }
                else { MobileData2File(); }

                MobileDataPath = MobileDataPath_0;
                MyCommond.CheckSimilarFile(ref MobileDataPath);

                if (paoTicketSaleLists == null) { paoTicketSaleLists = MyCommond.ToViewModel<PaoTicketSaleList>(ThisReceive, PaoDataPath, (new PaoTicketSaleList()).ExportTitle_zhTW(), 1); }

                if (ReadMobilePayData == null) { ReadMobilePayData = MyCommond.ToViewModel<MobilePay>(ThisReceive, MobileDataPath, (new MobilePay()).ExportTitle_zhTW(), 1); }
            }   //月
            else
            {
                //paoTicketSaleLists = FormData.ToList();
                MobileDataInput();
            }   //日

            List<CheckBoxs> _CB = CB.FindAll(x => x.Status == true).ToList();
            foreach (var item in _CB)
            {
                MyCommond.WriteLog(ThisReceive, $"創建 {item.Methode.ToString()} 站台 * 1");
                LrtClient client = new LrtClient(this, item.Methode);
                client.Start();
                LrtClients.Add(client);
            }
            IsRunning = true;
        }

        public void Close()
        {

        }

        #endregion

        #region 非同步線程執行程式

        public void Test()
        {
            MyCommond.WriteLog(ThisReceive, $"創建 測試 站台 * 1");
            LrtClient client = new LrtClient(this, LrtTxnClientSwitch.NA);
            //client.Start();
            LrtClients.Add(client);
        }

        private void MobileDataInput()
        {

        }

        #region 每月功能執行前整理 - 合併電支

        private void MobileData2File()
        {
            MyCommond.WriteLog(ThisReceive, $"開始電支檔案合併");
            MyCommond.WriteLog(ThisReceive, $"本次合併{OperationStartDate:yyyy年MM月份}的所有CSV檔");
            DateTime MonthFirstDay = Convert.ToDateTime(OperationStartDate.ToString("yyyy/MM/01 05:00:00"));
            DateTime MonthEndDay = MonthFirstDay.AddMonths(1);

            if (MonthFirstDay.Year != MonthEndDay.Year) { MonthEndDay = MonthEndDay.AddDays(1); } // 年底多匯入一天

            TreeForMobile root = null;
            string MobileDataMonth_FileName = MyCommond.Path_Program + MyCommond.Path_Analyze + $"MobileData_{MonthFirstDay.ToString("yyyy-MM")}.csv";

            MyCommond.WriteLog(ThisReceive, $"合併後檔案路徑及位置:{MobileDataMonth_FileName}");
            int RunCount = 0;
            int MaxCount = Convert.ToInt16((MonthEndDay - MonthFirstDay).ToString("dd"));

            List<MobilePay> MobilePays = new List<MobilePay>();

            string filenamepath = MyCommond.Path_Program + MyCommond.Path_Analyze;
            for (DateTime iDate = MonthFirstDay; iDate < MonthEndDay; iDate = iDate.AddDays(1))
            {
                MyCommond.WriteLog(ThisReceive, $"執行第{RunCount + 1}天電支合併");

                string SearchFileName;
                for (int i = 0; i < 2; i++)
                {
                    if (i == 0) SearchFileName = string.Format(FileName_QrCode, iDate.ToString("yyyy-MM-dd"));
                    else if (i == 1) SearchFileName = string.Format(FileName_CreditCard, iDate.ToString("yyyy-MM-dd"));
                    else break;
                    MyCommond.WriteLog(ThisReceive, $"本次提取的檔案:{SearchFileName}");
                    if (MyCommond.CheckFile(filenamepath + SearchFileName))
                    {
                        List<MobilePay> temp = MyCommond.ToViewModel<MobilePay>(ThisReceive, filenamepath + SearchFileName, new MobilePay().ExportTitle_zhTW(), 2);
                        MobilePays.AddRange(temp);
                        continue;
                        string[] CsvData = System.IO.File.ReadAllLines(filenamepath + SearchFileName);
                        foreach (var item in CsvData)
                        {
                            //WriteLog(1, MethodBase.GetCurrentMethod().Name.ToString(), $"提取資料:{item}");
                            string[] jtem = item.Split(',');
                            if (jtem[0] == "" || jtem[1] == "" || jtem[2] == "" || jtem[3] == "" || jtem.Length < 9) continue;
                            root = Insert(root, jtem[2], item);
                        }
                    }
                }
                if (RunCount > MaxCount + 10) break;
                RunCount++;
            }

            MyCommond.ExportData<MobilePay>(ThisReceive, new ExportConfig()
            {
                Path1 = filenamepath,
                DTime = new DateTime(),
                FirstName = "",
                FileName = "MobileData",
                SubName = MonthFirstDay.ToString("yyyy-MM"),
                OpLine = "",
                LastName = "csv",
            }, new MobilePay().ExportTitle_zhTW(), MobilePays, Encoding.UTF8);
            goto EndSub;

            if (root != null)
            {
                MyCommond.WriteLog(ThisReceive, $"準備匯出檔案");
                List<string> mergedData = new List<string>();
                Traverse(root, mergedData);
                File.WriteAllLines(MobileDataMonth_FileName, mergedData, Encoding.UTF8);
                MyCommond.WriteLog(ThisReceive, $"CSV files merged successfully!");
            }
            else
            {
                MyCommond.WriteLog(ThisReceive, $"No data found in the CSV files.");
            }

        EndSub:
            MyCommond.WriteLog(ThisReceive, $"電支月資料合併結束");
        }

        class TreeForMobile
        {
            public string SortData;
            public string StringLine;
            public TreeForMobile Left;
            public TreeForMobile Right;

            public TreeForMobile(string string1, string string2)
            {
                SortData = string1;
                StringLine = string2;
            }

        }

        private TreeForMobile Insert(TreeForMobile root, string SortData, string StringLine)
        {
            if (root == null)
            {
                MyCommond.WriteLog(ThisReceive, $"應該是第一次寫檔出現寫入：{StringLine}");
                return new TreeForMobile(SortData, StringLine);
            }
            else if (StringLine.CompareTo(root.StringLine) < 0)
            {
                MyCommond.WriteLog(ThisReceive, $"往左走寫入              ：{StringLine}");
                root.Left = Insert(root.Left, SortData, StringLine);
            }
            else if (StringLine.CompareTo(root.StringLine) > 0)
            {
                MyCommond.WriteLog(ThisReceive, $"往右走寫入              ：{StringLine}");
                root.Right = Insert(root.Right, SortData, StringLine);
            }
            else
            {
                MyCommond.WriteLog(ThisReceive, $"有了同樣的資料          ：{StringLine}");
            }
            return root;
        }

        private void Traverse(TreeForMobile root, List<string> mergedData)
        {
            if (root != null)
            {
                Traverse(root.Left, mergedData);
                mergedData.Add(root.StringLine);
                Traverse(root.Right, mergedData);
            }

        }

        #endregion

        #endregion

    }
}
