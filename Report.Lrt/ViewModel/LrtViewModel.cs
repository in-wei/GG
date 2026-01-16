using System;
//using System.Timers;
using System.Windows.Forms;
//--
using UsuallyCommond.MyEnum;


namespace Report.Lrt.ViewModel
{
    public class LrtViewModel
    {

    }

    public class PublicUse
    {
        public string Receive { get; set; }
        public ExecutionMode ExportMode { get; set; }
        public string Path { get; set; }
        public string Count { get; set; }
        public string Msg { get; set; }

        public override string ToString()
        {
            string exportString = "";

            exportString = $"{Receive},{ExportMode},{Path},{Count},{Msg}";

            return exportString;
        }

        public int ColumnCompare_zhTW(string GetTitle)
        {
            int ColumnValue;
            switch (GetTitle)
            {
                case "0": ColumnValue = 0; break;
                case "1": ColumnValue = 1; break;
                case "2": ColumnValue = 2; break;
                case "3": ColumnValue = 3; break;
                case "4": ColumnValue = 4; break;
                default: ColumnValue = -1; break;
            }
            return ColumnValue;
        }

    }

    public class SqlServer
    {
        public string IP { get; set; }
        public string Calalog { get; set; }
        public string ID { get; set; }
        public string Password { get; set; }

        public override string ToString()
        {
            string exportString = "";

            exportString = $"{IP},{Calalog},{ID},{Password}";

            return exportString;
        }

        public int ColumnCompare_zhTW(string GetTitle)
        {
            int ColumnValue;
            switch (GetTitle)
            {
                case "地址": ColumnValue = 0; break;
                case "資料庫": ColumnValue = 1; break;
                case "使用者": ColumnValue = 2; break;
                case "密碼": ColumnValue = 3; break;
                default: ColumnValue = -1; break;
            }
            return ColumnValue;
        }

    }

    public class CheckBoxs
    {
        public LrtTxnClientSwitch Methode { get; set; }
        public bool Status { get; set; }
        public CheckBox Item { get; set; }
        
        public override string ToString()
        {
            string exportString = "";

            exportString = $"{Methode},{Status},{Item}";

            return exportString;
        }

        public int ColumnCompare_zhTW(string GetTitle)
        {
            int ColumnValue;
            switch (GetTitle)
            {
                case "線程": ColumnValue = 0; break;
                case "狀態": ColumnValue = 1; break;
                case "物件": ColumnValue = 2; break;
                default: ColumnValue = -1; break;
            }
            return ColumnValue;
        }

    }

    public class Timers
    {
        public LrtTxnClientSwitch Methode { get; set; }
        public StatusEnum Status { get; set; }
        public Timer Item { get; set; }

        public override string ToString()
        {
            string exportString = "";

            exportString = $"{Methode},{Status},{Item}";

            return exportString;
        }

        public int ColumnCompare_zhTW(string GetTitle)
        {
            int ColumnValue;
            switch (GetTitle)
            {
                case "線程": ColumnValue = 0; break;
                case "狀態": ColumnValue = 1; break;
                case "物件": ColumnValue = 2; break;
                default: ColumnValue = -1; break;
            }
            return ColumnValue;
        }

    }

    public class ProgressBars
    {
        public LrtTxnClientSwitch Methode { get; set; }
        public ProgressBar Item { get; set; }

        public override string ToString()
        {
            string exportString = "";

            exportString = $"{Methode},{Item}";

            return exportString;
        }

        public int ColumnCompare_zhTW(string GetTitle)
        {
            int ColumnValue;
            switch (GetTitle)
            {
                case "線程": ColumnValue = 0; break;
                case "物件": ColumnValue = 1; break;
                default: ColumnValue = -1; break;
            }
            return ColumnValue;
        }

    }

    #region 自有讀檔資料

    public class EachSellReport
    {
        public string Operation { get; set; }
        public string Station { get; set; }
        #region +
        public string Ticket_OneDay { get; set; }
        public string Ticket_10 { get; set; }
        public string Ticket_12 { get; set; }
        public string Ticket_15 { get; set; }
        public string Ticket_20 { get; set; }
        public string Ticket_25 { get; set; }
        public string Ticket_30 { get; set; }
        public string Ticket_bike { get; set; }
        public string OpenTrafficSetTicket { get; set; }
        public string TPass { get; set; }
        public string Ticket_Group { get; set; }
        public string AfcTackOut { get; set; }
        public string Repair { get; set; }
        public string OtherPlus { get; set; }
        #endregion

        #region -
        public string ConcessionRefund { get; set; }
        public string RefundNotice { get; set; }
        public string VavmRefund { get; set; }
        public string PvRefund { get; set; }
        public string CreditCardRefund { get; set; }
        public string QrCodeRefund { get; set; }
        public string OperationInterrupt { get; set; }
        public string SaleCardRefund { get; set; }
        public string GroupRefund { get; set; }
        public string OtherRefund { get; set; }
        #endregion

        public override string ToString()
        {
            return $"" +
                $"{Operation}," +
                $"{Station}," +
                $"{Ticket_OneDay}," +
                $"{Ticket_10}," +
                $"{Ticket_12}," +
                $"{Ticket_15}," +
                $"{Ticket_20}," +
                $"{Ticket_25}," +
                $"{Ticket_30}," +
                $"{Ticket_bike}," +
                $"{OpenTrafficSetTicket}," +
                $"{TPass}," +
                $"{Ticket_Group}," +
                $"{AfcTackOut}," +
                $"{Repair}," +
                $"{OtherPlus}," +
                $"{ConcessionRefund}," +
                $"{RefundNotice}," +
                $"{VavmRefund}," +
                $"{PvRefund}," +
                $"{CreditCardRefund}," +
                $"{QrCodeRefund}," +
                $"{OperationInterrupt}," +
                $"{SaleCardRefund}," +
                $"{GroupRefund}," +
                $"{OtherRefund}";
        }

        public string ExportTitle_English()
        {
            return $"" +
                $"Operation," +
                $"Station," +
                $"Ticket_OneDay," +
                $"Ticket_10," +
                $"Ticket_12," +
                $"Ticket_15," +
                $"Ticket_20," +
                $"Ticket_25," +
                $"Ticket_30," +
                $"Ticket_bike," +
                $"OpenTrafficSetTicket," +
                $"TPass," +
                $"Ticket_Group," +
                $"AfcTackOut," +
                $"Repair," +
                $"OtherPlus," +
                $"ConcessionRefund," +
                $"RefundNotice," +
                $"VavmRefund," +
                $"PvRefund," +
                $"CreditCardRefund," +
                $"QrCodeRefund," +
                $"OperationInterrupt," +
                $"SaleCardRefund," +
                $"GroupRefund," +
                $"OtherRefund";
        }

        public string ExportTitle_zhTW()
        {
            return $"" +
                $"營運," +
                $"車站," +
                $"一日票," +
                $"優待票(10元)," +
                $"優待票(12元)," +
                $"優待票(15元)," +
                $"單程票(20元)," +
                $"單程票(25元)," +
                $"單程票(30元)," +
                $"自行車單程票," +
                $"通車週年套票," +
                $"TPASS售卡收入," +
                $"團體票," +
                $"AFC設備異常紀錄單(VAVM取出)," +
                $"補繳車資單收入," +
                $"其他收入," +
                $"優待票代墊款," +
                $"旅客退費通知單退費," +
                $"VAVM設備異常退費," +
                $"誤刷卡退費," +
                $"行動支付退費(信用卡)," +
                $"行動支付退費(乘車碼)," +
                $"營運中斷退費," +
                $"加值誤按售卡退費," +
                $"團體票退費," +
                $"其他";
        }

    }

    public class PaoTicketSaleList
    {
        public string Operation { get; set; }
        public string Station { get; set; }
        #region 收入
        public string Ticket_OneDay { get; set; }
        public string Ticket10 { get; set; }
        public string Ticket12 { get; set; }
        public string Ticket15 { get; set; }
        public string Ticket20 { get; set; }
        public string Ticket25 { get; set; }
        public string Ticket30 { get; set; }
        public string TicketBike { get; set; }
        public string OpenTrafficSetTicket { get; set; }
        public string TPASS { get; set; }
        public string GroupCount { get; set; }
        public string GroupPeople { get; set; }
        public string GroupAmt { get; set; }
        public string Lanwair1 { get; set; }
        public string Lanwair2 { get; set; }
        public string AfcTakeOut { get; set; }
        public string RepairTicket { get; set; }
        #endregion
        # region 其他
        public string ExchangeTicket20 { get; set; }
        public string ExchangeTicket25 { get; set; }
        public string ExchangeTicket30 { get; set; }
        public string KKDay { get; set; }
        public string Ntpc { get; set; }
        public string SplitTicker { get; set; }
        public string ExchangeOneDay { get; set; }
        #endregion
        # region 支出
        public string ConcessionRefund { get; set; }
        public string RefundNotice { get; set; }
        public string VavmRefund { get; set; }
        public string PvRefund { get; set; }
        public string CreditCardRefund { get; set; }
        public string QrCodeRefund { get; set; }
        public string SaleCardRefund { get; set; }
        public string GroupRefund { get; set; }
        #endregion
        
        public override string ToString()
        {
            return $"{Operation}," +
                $"{Station}," +
                $"" +
                $"{Ticket_OneDay}," +
                $"{Ticket10},{Ticket12},{Ticket15}," +
                $"{Ticket20},{Ticket25},{Ticket30}," +
                $"{TicketBike},{OpenTrafficSetTicket},{TPASS}," +
                $"{GroupCount},{GroupPeople},{GroupAmt}," +
                $"{Lanwair1},{Lanwair2}," +
                $"{AfcTakeOut},{RepairTicket}," +
                $"" +
                $"{ExchangeTicket20},{ExchangeTicket25},{ExchangeTicket30}," +
                $"{KKDay},{Ntpc},{SplitTicker},{ExchangeOneDay}," +
                $"" +
                $"{ConcessionRefund},{RefundNotice},{VavmRefund}," +
                $"{PvRefund},{CreditCardRefund},{QrCodeRefund}," +
                $"{SaleCardRefund},{GroupRefund}";
        }

        public string ExportTitle_zhTW()
        {
            ///營運,
            ///車站,
            ///一日票,
            ///優待票10元,優待票12元,優待票15元,
            ///單程票20元,單程票25元,單程票30元,
            ///自行車票,通車紀念套票,行政院通勤月票紀念卡,
            ///團體票張數,團體票人數,團體票金額,
            ///活動1,活動2,
            ///設備取出,補票,
            ///兌換單程票20元,兌換單程票25元,兌換單程票30元,
            ///一日票(KKDay),一日票(新北幣),切票,兌換一日票,
            ///優待票代墊款,退費通知,自動售票機退費,
            ///刷卡機退費,信用卡退費,乘車碼退費,
            ///退卡,團體票退費

            return $"營運," +
                $"車站," +
                $"" +
                $"一日票," +
                $"優待票10元,優待票12元,優待票15元," +
                $"單程票20元,單程票25元,單程票30元," +
                $"自行車票,通車紀念套票,行政院通勤月票紀念卡," +
                $"團體票張數,團體票人數,團體票金額," +
                $"活動1,活動2," +
                $"設備取出,補票," +
                $"" +
                $"兌換單程票20元,兌換單程票25元,兌換單程票30元," +
                $"一日票(KKDay),一日票(新北幣),切票,兌換一日票," +
                $"" +
                $"優待票代墊款,退費通知,自動售票機退費," +
                $"刷卡機退費,信用卡退費,乘車碼退費," +
                $"退卡,團體票退費";
        }
        
        public string ExportTitle_English()
        {
            return $"OperationDate," +
                $"Station," +
                $"" +
                $"Ticket_OneDay," +
                $"Ticket10,Ticket12,Ticket15," +
                $"Ticket20,Ticket25,Ticket30," +
                $"TicketBike,OpenTrafficSetTicket,TPASS" +
                $"GroupCount,GroupPeople,GroupAmt," +
                $"Lanwair1,Ticket_Lanwair," +
                $"AfcTakeOut,RepairTicket," +
                $"" +
                $"ExchangeTicket20,ExchangeTicket25,ExchangeTicket30," +
                $"KKDay,Ntpc,SplitTicker,ExchangeOneDay," +
                $"" +
                $"ConcessionRefund,RefundNotice,VavmRefund," +
                $"PvRefund,CreditCardRefund,QrCodeRefund," +
                $"SaleCardRefund,GroupRefund";
        }

    }

    public class LrtPaoService
    {
        public string Id { get; set; }
        public LrtPaoServiceList Txn { get; set; }
        public bool TF { get; set; }

        public string ExportTitle_zhTW()
        {
            return $"身分字串,交易類型,是否";
        }

        public string ExportTitle_English()
        {
            return $"Id,Txn,TF";
        }

        public override string ToString()
        {
            return $"{Id},{Txn},{TF}";
        }

    }

    public class LrtStationList
    {
        public string OperationLine { get; set; }
        public string CodeNumber { get; set; }
        public string CodeName { get; set; }
        public string StationName { get; set; }
        public string StationPao { get; set; }

        public string ExportTitle_zhTW()
        {
            return $"路線,代碼,代號,名稱,詢問處";
        }

        public string ExportTitle_English()
        {
            return $"OperationLine,CodeNumber,CodeName,StationName,StationPao";
        }

        public override string ToString()
        {
            return $"{OperationLine},{CodeNumber},{CodeName},{StationName},{StationPao}";
        }

    }

    public class GroupTicketList
    {
        public string SaleDate { get; set; }
        public string SaleStation { get; set; }
        public string Entry { get; set; }
        public string Exit { get; set; }
        public string People { get; set; }
        public string Amount { get; set; }
        public string TotleAmount { get; set; }
        public string SaleTime { get; set; }
        public string UseDate { get; set; }
        public string SaleWay { get; set; }
        public string Remark { get; set; }

        public string ExportTitle_zhTW()
        {
            return $"" +
                $"販售日期," +
                $"販售車站," +
                $"起," +
                $"迄," +
                $"人數," +
                $"單價," +
                $"總價," +
                $"銷售時間," +
                $"使用日期," +
                $"開立方式," +
                $"備註";
        }

        public string ExportTitle_English()
        {
            return $"" +
                $"SaleDate," +
                $"SaleStation," +
                $"Entry," +
                $"Exit," +
                $"People," +
                $"Amount," +
                $"TotleAmount," +
                $"SaleTime," +
                $"UseDate," +
                $"SaleWay," +
                $"Remark";
        }

        public override string ToString()
        {
            return $"" +
                $"{SaleDate}," +
                $"{SaleStation}," +
                $"{Entry}," +
                $"{Exit}," +
                $"{People}," +
                $"{Amount}," +
                $"{TotleAmount}," +
                $"{SaleTime}," +
                $"{UseDate}," +
                $"{SaleWay}," +
                $"{Remark}";
        }

    }

    public class GroupTicketList_Old
    {
        public string SaleDate { get; set; }
        public string SaleStation { get; set; }
        public string EntryStation { get; set; }
        public string ExitStation { get; set; }
        public string People { get; set; }
        public string UsingTime { get; set; }
        public string UsingDate { get; set; }

        public string ExportTitle_zhTW()
        {
            return $"報表中文名,報表英文名," +
                $"報表日期欄,報表日期行," +
                $"起始行,起始欄," +
                $"結束行,結束欄," +
                $"斷點行,斷點行數," +
                $"斷點欄,斷點欄數";
        }

        public string ExportTitle_English()
        {
            return $"ReportNameCh,ReportNameEng," +
                $"ReportDateRow,ReportDateColumn," +
                $"StartRow,StartColumn," +
                $"EndRow,EndColumn," +
                $"BreakOffRow,BreakOffRowNum," +
                $"BreakOffColumn,BreakOffColumnNum";
        }

        public override string ToString()
        {
            return $"{SaleDate},{SaleStation}," +
                $"{EntryStation},{ExitStation}," +
                $"{People},{UsingTime}," +
                $"{UsingDate}";
        }

    }

    public class ReportFileTemp
    {
        public string ReportNameCh { get; set; }
        public string ReportNameEng { get; set; }
        public string ReportDateColumn { get; set; }
        public string ReportDateRow { get; set; }
        public string StartColumn { get; set; }
        public string StartRow { get; set; }
        public string EndColumn { get; set; }
        public string EndRow { get; set; }
        public string BreakOffRow { get; set; }
        public string BreakOffRowNum { get; set; }
        public string BreakOffColumn { get; set; }
        public string BreakOffColumnNum { get; set; }
        
        public string ExportTitle_zhTW()
        {
            return $"報表中文名,報表英文名," +
                $"產製日期欄(英),產製日期行(數)," +
                $"起始欄(英),起始行(數)," +
                $"結束欄(英),結束行(數)," +
                $"斷點行(數),斷點行數," +
                $"斷點欄(英),斷點欄數";
        }

        public string ExportTitle_English()
        {
            return $"ReportNameCh,ReportNameEng," +
                $"ReportDateColumn,ReportDateRow," +
                $"StartColumn,StartRow," +
                $"EndColumn,EndRow," +
                $"BreakOffRow,BreakOffRowNum," +
                $"BreakOffColumn,BreakOffColumnNum";
        }

        public override string ToString()
        {
            return $"{ReportNameCh},{ReportNameEng}," +
                $"{StartColumn},{ReportDateRow}," +
                $"{StartColumn},{StartRow}," +
                $"{EndColumn},{EndRow}," +
                $"{BreakOffRow},{BreakOffRowNum}," +
                $"{BreakOffColumn},{BreakOffColumnNum}";
        }

    }

    public class ReportFileTemp2
    {
        public string OperationLine { get; set; }
        public string ReportNameCh { get; set; }
        public string ReportNameEng { get; set; }
        public string ReportDateColumn { get; set; }
        public string ReportDateRow { get; set; }
        public string StartColumn { get; set; }
        public string StartRow { get; set; }
        public string EndColumn { get; set; }
        public string EndRow { get; set; }
        public string BreakOffRow { get; set; }
        public string BreakOffRowNum { get; set; }
        public string BreakOffColumn { get; set; }
        public string BreakOffColumnNum { get; set; }
        public string MobilePayColumn { get; set; }
        public string OwnTiketColumn { get; set; }

        public string ExportTitle_zhTW()
        {
            return $"路線," +
                $"報表中文名,報表英文名," +
                $"產製日期欄(英),產製日期行(數)," +
                $"起始欄(英),起始行(數)," +
                $"結束欄(英),結束行(數)," +
                $"斷點行(數),斷點行數," +
                $"斷點欄(英),斷點欄數," +
                $"電子支付起始欄(英),自有票起始欄(英),";
        }

        public string ExportTitle_English()
        {
            return $"OperationLine," +
                $"ReportNameCh,ReportNameEng," +
                $"ReportDateColumn,ReportDateRow," +
                $"StartColumn,StartRow," +
                $"EndColumn,EndRow," +
                $"BreakOffRow,BreakOffRowNum," +
                $"BreakOffColumn,BreakOffColumnNum," +
                $"MobilePayColumn,OwnTiketColumn,";
        }

        public override string ToString()
        {
            return $"{OperationLine}," +
                $"{ReportNameCh},{ReportNameEng}," +
                $"{StartColumn},{ReportDateRow}," +
                $"{StartColumn},{StartRow}," +
                $"{EndColumn},{EndRow}," +
                $"{BreakOffRow},{BreakOffRowNum}," +
                $"{BreakOffColumn},{BreakOffColumnNum}," +
                $"{MobilePayColumn},{OwnTiketColumn},";
        }

    }

    public class Subsidy
    {
        public string Operation_Area { get; set; }
        public string Operation_Line { get; set; }
        public string StartStation { get; set; }
        public string EndStation { get; set; }
        public string ShoudPay { get; set; }
        public string OdAmt { get; set; }

        public string ExportTitle_English()
        {
            string result;

            result = $"" +
                $"Operation_Area," +
                $"Operation_Line," +
                $"StartStation," +
                $"EndStation," +
                $"ShoudPay," +
                $"OdAmt";

            return result;
        }

        public string ExportTitle_zhTW()
        {
            string result;

            result = $"" +
                $"營運區域," +
                $"路線代號," +
                $"起站," +
                $"訖站," +
                $"運價," +
                $"差額";

            return result;
        }

        public override string ToString()
        {
            string result;

            result = $"" +
                $"{Operation_Area}," +
                $"{Operation_Line}," +
                $"{StartStation}," +
                $"{EndStation}," +
                $"{ShoudPay}," +
                $"{OdAmt}";

            return result;
        }

    }

    public class OdSubsidy
    {
        public string EntryStation { get; set; }
        public string ExitStation { get; set; }
        public double Num { get; set; }

        public string ExportTitle_zhTW()
        {
            return $"起站," +
                $"訖站," +
                $"數量";
        }

        public string ExportTitle_English()
        {
            return $"EntryStation," +
                $"ExitStation," +
                $"Num";
        }

        public override string ToString()
        {
            return $"{EntryStation}," +
                $"{ExitStation}," +
                $"{Num}";
        }

    }

    public class ReSql
    {
        public string Using { get; set; }
        public bool UpData { get; set; }

        public string ExportTitle_zhTW()
        {
            return $"使用語法,是否更新";
        }

        public string ExportTitle_English()
        {
            return $"Using,UpData";
        }

        public override string ToString()
        {
            return $"{Using},{UpData}";
        }

    }

    public class SqlList
    {
        public string UseNameCh { get; set; }
        public string FileName1 { get; set; }
        public string FileName2 { get; set; }
        public string Remark { get; set; }
        public string Using { get; set; }

        public string ExportTitle_zhTW()
        {
            return $"報表中文,語法名稱1,語法名稱2,停用標記,使用語法";
        }

        public string ExportTitle_English()
        {
            return $"UseNameCh,FileName1,FileName2,Remark,Using";
        }

        public override string ToString()
        {
            return $"{UseNameCh},{FileName1},{FileName2},{Remark},{Using}";
        }

    }

    public class MobileCardType
    {
        public string CardName { get; set; }
        public string CardType { get; set; }
        public string CardSubType { get; set; }

        public string ExportTitle_zhTW()
        {
            return $"卡片名稱,卡種,卡別";
        }

        public string ExportTitle_English()
        {
            return $"CardName,CardType,CardSubType";
        }

        public override string ToString()
        {
            return $"{CardName},{CardType},{CardSubType}";
        }

    }

    public class ReportType1
    {
        public const string Day = "Day";
        public const string Station = "Station";
    }

    public class mC1
    {
        public int General { get; set; }
        public int Concession { get; set; }
        public int Bike { get; set; }
        public int OneTicketCount1 { get; set; }
        public int OneTicketTimes1 { get; set; }
        public int OneTicketCount2 { get; set; }
        public int OneTicketTimes2 { get; set; }
        public int GroupCount { get; set; }
        public int GroupPeople { get; set; }
        
    }

    #endregion

    #region 行動支付讀取用
    public class MobilePay
    {
        public string CardCN { set; get; }  //0
        public string CardSubType { set; get; }    //1
        public string OpeLine { set; get; }
        public DateTime EntryTime { set; get; }     //2
        public string EntryStation { set; get; }    //3
        public DateTime ExitTime { set; get; }  //4
        public string ExitStation { set; get; } //5
        public int ShouldPay { set; get; }  //6
        public int Fine { set; get; }   //7
        public string Volume { set; get; }  //8
        public int RealPay { set; get; }    //9
        public int Subsidization { set; get; }  //10
        public int PayBack { set; get; }    //11
        public DateTime PayBackTime { set; get; }   //12
        
        public string ExportTitle_zhTW()
        {
            return $"卡號," +
                $"卡別," +
                $"營運路線," +
                $"進站時間," +
                $"進站別," +
                $"出站時間," +
                $"出站別," +
                $"應收金額," +
                $"罰款," +
                $"交易類別," +
                $"實收金額," +
                $"優惠金額," +
                $"退款," +
                $"退款時間";
        }

        public string ExportTitle_English()
        {
            return $"" +
                $"CardCN," +
                $"CardSubType," +
                $"OpeLine," +
                $"EntryTime," +
                $"EntryStation," +
                $"ExitTime," +
                $"ExitStation," +
                $"ShouldPay," +
                $"Fine," +
                $"Volume," +
                $"RealPay," +
                $"Subsidization," +
                $"PayBack," +
                $"PayBackTime";
        }

        public override string ToString()
        {
            return $"" +
                $"{CardCN}," +
                $"{CardSubType}," +
                $"{OpeLine}," +
                $"{EntryTime}," +
                $"{EntryStation}," +
                $"{ExitTime}," +
                $"{ExitStation}," +
                $"{ShouldPay}," +
                $"{Fine}," +
                $"{Volume}," +
                $"{RealPay}," +
                $"{Subsidization}," +
                $"{((PayBackTime == new DateTime()) ? "" : $"{PayBack}")}," +
                $"{((PayBackTime == new DateTime()) ? "" : $"{PayBackTime}")}";
        }

    }

    public class MobilePayVolum
    {
        public const string Normal = "票價扣款";
        public const string NoEnrty = "進站補扣款";       //出站時產生  (沒刷進站刷出站)
        public const string NoExit = "出站補扣款";       //半夜結帳產生(有刷進站沒出站)
        public const string OutTime = "逾時罰款";         //進站後很久才出站
        public const string InTime = "同站進出時間內不扣款";  //進站後在很短的時間內刷出站
        public const string NextDayExit = "跨日出站扣款";     //同正常進出，只是跨日了
    }

    public class MobilePayTimeStation
    {
        public string Station { get; set; }
        public int Counts { get; set; }

        public string ExportTitle_zhTW()
        {
            return $"車站,數量";
        }

        public string ExportTitle_English()
        {
            return $"Station,Counts";
        }

        public override string ToString()
        {
            return $"{Station},{Counts}";
        }

    }

    public class MobileReportTicketForStation
    {
        public string Timer { get; set; }
        public string Station { get; set; }
        public string CardType { get; set; }
        public int Counts { get; set; }

        public string ExportTitle_zhTW()
        {
            return $"時間,車站,卡種,數量";
        }

        public string ExportTitle_English()
        {
            return $"Timer,Station,CardType,Counts";
        }

        public override string ToString()
        {
            return $"{Timer},{Station},{CardType},{Counts}";
        }

    }

    public class MobileReportAmtForStation
    {
        public string Station { get; set; }
        public string CardType { get; set; }
        public int VolumnAmt { get; set; }
        public int FineAmt { get; set; }

        public string ExportTitle_zhTW()
        {
            return $"車站,卡種,車資,罰款";
        }

        public string ExportTitle_English()
        {
            return $"Station,CardType,VolumnAmt,FineAmt";
        }

        public override string ToString()
        {
            return $"{Station},{CardType},{VolumnAmt},{FineAmt}";
        }

    }

    public class MobileReportTicketForDay
    {
        public string rDay { get; set; }
        public string CardType { get; set; }
        public int Counts { get; set; }

        public string ExportTitle_zhTW()
        {
            return $"日期,卡種,數量,";
        }

        public string ExportTitle_English()
        {
            return $"rDay,CardType,Counts,";
        }

        public override string ToString()
        {
            return $"{rDay},{CardType},{Counts}";
        }

    }

    public class MobileListAmtForStation
    {
        public string Station { get; set; }
        public string CardKind { get; set; }
        public int RealPay { get; set; }

        public string ExportTitle_zhTW()
        {
            return $"車站,卡種,實際支付,";
        }

        public string ExportTitle_English()
        {
            return $"Station,CardKind,RealPay,";
        }

        public override string ToString()
        {
            return $"{Station},{CardKind},{RealPay}";
        }

    }

    public class MobileReportAmtForStation1
    {
        public string Station { set; get; }
        public int SumAmt { get; set; }

        public string ExportTitle_zhTW()
        {
            return $"車站,金額";
        }

        public string ExportTitle_English()
        {
            return $"Station,SumAmt";
        }

        public override string ToString()
        {
            return $"{Station},{SumAmt}";
        }

    }

    public class MobileListAmtForStation2
    {
        public string Station { get; set; }
        public int CreditAmt { get; set; }
        public int QrAmt { get; set; }
        
        public string ExportTitle_zhTW()
        {
            return $"車站,信用卡,乘車碼,";
        }

        public string ExportTitle_English()
        {
            return $"Station,CreditAmt,QrAmt,";
        }

        public override string ToString()
        {
            return $"{Station},{CreditAmt},{QrAmt}";
        }

    }

    #endregion

    #region 報表

    #region 營運資料統計_運量

    public class DayVolume
    {
        #region T
        // public string DateTime { get; set; }                    //日期
        //public string Total { get; set; }                       //總營收
        public string Electron_Exit_num { get; set; }           //普通卡
        public string Student_Exit_num { get; set; }           //學生
        public string Welfare_Exit_num { get; set; }            //社福卡
        public string All_Pass_Common_Exit_num { get; set; }    //1280普
        public string All_Pass_Student_Exit_num { get; set; }   //1280學
        public string Credit_num { get; set; }                  //信用卡
        public string Qrcode_num { get; set; }                  //QRCode
        public string SOneTkt_num { get; set; }                 //單程票
        public string SConcessionTkt_num { get; set; }          //單程優待票
        public string SBikeTkt_num { get; set; }                //自行車票
        public string OneDay1_num { get; set; }                 //一日票(公版)
        public string OneDay_Exit_num { get; set; }             //一日票(電票)
        public string Lanwair1_num { get; set; }                 //活動1
        public string Lanwair2_num { get; set; }                 //活動2
                                                                 //public string Lanwair3_num { get; set; }                 //活動3
                                                                 //public string Lanwair4_num { get; set; }                 //活動4
        public string Group_Tik_num { get; set; }               //團體票張數
        public string Group_Peo_num { get; set; }               //團體票人數
        public string Office_Staff_num { get; set; }            //員工卡
                                                                //public string Lanwair_num { get; set; }                 //活動
                                                                //public string Lanwair5_num { get; set; }                 //活動5
        #endregion

        public int ColumnCompare_zhTW(string GetTitle)
        {
            int ColumnValue;
            switch (GetTitle)
            {
                case "普通卡":       ColumnValue =  0; break;
                case "學生卡":       ColumnValue =  1; break;
                case "優待卡":       ColumnValue =  2; break;
                case "一般交通月票": ColumnValue =  3; break;
                case "學生交通月票": ColumnValue =  4; break;
                case "信用卡":       ColumnValue =  5; break;
                case "乘車碼":       ColumnValue =  6; break;
                case "一般單程票":   ColumnValue =  7; break;
                case "優待單程票":   ColumnValue =  8; break;
                case "自行車單程票": ColumnValue =  9; break;
                case "一日票":       ColumnValue = 10; break;
                case "一日電票":     ColumnValue = 11; break;
                case "活動1":        ColumnValue = 12; break;
                case "活動2":        ColumnValue = 13; break;
                case "團體票張數":   ColumnValue = 14; break;          
                case "團體票人數":   ColumnValue = 15; break;          
                case "員工卡":       ColumnValue = 16; break;   
                default: ColumnValue = -1; break;
            }
            return ColumnValue;
        }

        public int ColumnCompare_English(string GetTitle)
        {
            int ColumnValue;
            switch (GetTitle)
            {
                case "Electron_Exit_num":         ColumnValue =  0; break;
                case "Student_Exit_num":          ColumnValue =  1; break;
                case "Welfare_Exit_num":          ColumnValue =  2; break;
                case "All_Pass_Common_Exit_num":  ColumnValue =  3; break;
                case "All_Pass_Student_Exit_num": ColumnValue =  4; break;
                case "Credit_num":                ColumnValue =  5; break;
                case "Qrcode_num":                ColumnValue =  6; break;
                case "SOneTkt_num":               ColumnValue =  7; break;
                case "SConcessionTkt_num":        ColumnValue =  8; break;
                case "SBikeTkt_num":              ColumnValue =  9; break;
                case "OneDay1_num":               ColumnValue = 10; break;
                case "OneDay_Exit_num":           ColumnValue = 11; break;
                case "Lanwair1_num":              ColumnValue = 12; break;
                case "Lanwair2_num":              ColumnValue = 13; break;
                case "Group_Tik_num":             ColumnValue = 14; break;
                case "Group_Peo_num":             ColumnValue = 15; break;
                case "Office_Staff_num":          ColumnValue = 16; break;
                default: ColumnValue = -1; break;
            }
            return ColumnValue;
        }

        public string ExportTitle_zhTW()
        {
            return $"普通卡,學生卡,優待卡," +
                $"一般交通月票,學生交通月票,信用卡," +
                $"乘車碼,一般單程票,優待單程票," +
                $"自行車單程票,一日票,一日電票," +
                $"活動1,活動2,團體票張數," +
                $"團體票人數,員工卡";
        }

        public string ExportTitle_English()
        {
            return $"Electron_Exit_num,Student_Exit_num,Welfare_Exit_num," +
                $"All_Pass_Common_Exit_num,All_Pass_Student_Exit_num,Credit_num," +
                $"Qrcode_num,SOneTkt_num,SConcessionTkt_num," +
                $"SBikeTkt_num,OneDay1_num,OneDay_Exit_num" +
                $"Lanwair1_num,Lanwair2_num,Group_Tik_num," +
                $"Group_Peo_num,Office_Staff_num";
        }

        public override string ToString()
        {
            return $"{Electron_Exit_num},{Student_Exit_num},{Welfare_Exit_num}," +
                $"{All_Pass_Common_Exit_num},{All_Pass_Student_Exit_num},{Credit_num}," +
                $"{Qrcode_num},{SOneTkt_num},{SConcessionTkt_num}," +
                $"{SBikeTkt_num},{OneDay1_num},{OneDay_Exit_num}," +
                $"{Lanwair1_num},{Lanwair2_num},{Group_Tik_num}," +
                $"{Group_Peo_num},{Office_Staff_num}";
        }

    }

    #endregion

    #region 營運資料統計_營收

    public class DayAmount
    {
        #region T
        // public string DateTime { get; set; }                    //日期
        //public string Total { get; set; }                       //總營收
        public string Electront_Amt { get; set; }       
        public string Student_Amt { get; set; }         
        public string Welfare_Amt { get; set; }         
        public string Credit_Amt { get; set; }          
        public string Qrcode_Amt { get; set; }          
        public string SOneTkt_Amt { get; set; }         
        public string SConcessionTkt_Amt { get; set; }  
        public string SBikeTkt_Amt { get; set; }        
        public string OneDay1_Amt { get; set; }         
        public string OneDay_Exit_Amt { get; set; }     
        public string Lanwair1_Amt { get; set; }        
        public string Lanwair2_Amt { get; set; }        
                                                        
        public string Group_Amt { get; set; }           
        public string Lanwair4_Amt { get; set; }        
        public string Retik_Amt { get; set; }           
                                                        
        public string ConcessionTicket { get; set; }    
        public string ConcessionPoint { get; set; }     
        public string SOneConcessionReBack { get; set; }
        public string YouBickToLrt { get; set; }        
        public string AllPassPayback { get; set; }      
        public string RideSubsidy { get; set; }
        #endregion

        public int ColumnCompare_zhTW(string GetTitle)
        {
            int ColumnValue;
            switch (GetTitle)
            {
                case "普通票":          ColumnValue =  0; break;
                case "學生票":          ColumnValue =  1; break;
                case "優待票":          ColumnValue =  2; break;
                case "信用卡":          ColumnValue =  3; break;
                case "乘車碼":          ColumnValue =  4; break;
                case "一般單程票":      ColumnValue =  5; break;
                case "優待單程票":      ColumnValue =  6; break;
                case "自行車單程票":    ColumnValue =  7; break;
                case "一日票":          ColumnValue =  8; break;
                case "一日電票":        ColumnValue =  9; break;
                case "活動1":           ColumnValue = 10; break;
                case "活動2":           ColumnValue = 11; break;
                case "團體票":          ColumnValue = 12; break;
                case "活動4":           ColumnValue = 13; break;
                case "補票":            ColumnValue = 14; break;
                case "社福票":          ColumnValue = 15; break;
                case "社福點數":        ColumnValue = 16; break;
                case "社福代墊款":      ColumnValue = 17; break;
                case "微笑單車轉乘":    ColumnValue = 18; break;
                case "交通月票":        ColumnValue = 19; break;
                case "里程補貼":        ColumnValue = 20; break;
                default: ColumnValue = -1; break;
            }
            return ColumnValue;
        }

        public int ColumnCompare_English(string GetTitle)
        {
            int ColumnValue;
            switch (GetTitle)
            {
                case "Electront_Amt":           ColumnValue =  0; break;
                case "Student_Amt":             ColumnValue =  1; break;
                case "Welfare_Amt":             ColumnValue =  2; break;
                case "Credit_Amt":              ColumnValue =  3; break;
                case "Qrcode_Amt":              ColumnValue =  4; break;
                case "SOneTkt_Amt":             ColumnValue =  5; break;
                case "SConcessionTkt_Amt":      ColumnValue =  6; break;
                case "SBikeTkt_Amt":            ColumnValue =  7; break;
                case "OneDay1_Amt":             ColumnValue =  8; break;
                case "OneDay_Exit_Amt":         ColumnValue =  9; break;
                case "Lanwair1_Amt":            ColumnValue = 10; break;
                case "Lanwair2_Amt":            ColumnValue = 11; break;
                case "Group_Amt":               ColumnValue = 12; break;
                case "Lanwair4_Amt":            ColumnValue = 13; break;
                case "Retik_Amt":               ColumnValue = 14; break;
                case "ConcessionTicket":        ColumnValue = 15; break;
                case "ConcessionPoint":         ColumnValue = 16; break;
                case "SOneConcessionReBack":    ColumnValue = 17; break;
                case "YouBickToLrt":            ColumnValue = 18; break;
                case "AllPassPayback":          ColumnValue = 19; break;
                case "RideSubsidy":             ColumnValue = 20; break;
                default: ColumnValue = -1; break;
            }
            return ColumnValue;
        }

        public string ExportTitle_zhTW()
        {
            return $"普通票,學生票,優待票," +
                $"信用卡,乘車碼,一般單程票," +
                $"優待單程票,自行車單程票,一日票," +
                $"一日電票,活動1,活動2" +
                $"團體票,活動4,補票," +
                $"社福票,社福點數," +
                $"社福代墊款,微笑單車轉乘," +
                $"交通月票,里程補貼";
        }

        public string ExportTitle_English()
        {
            return $"Electront_Amt,Student_Amt,Welfare_Amt," +
                $"Credit_Amt,Qrcode_Amt,SOneTkt_Amt," +
                $"SConcessionTkt_Amt,SBikeTkt_Amt,OneDay1_Amt," +
                $"OneDay_Exit_Amt,Lanwair1_Amt,Lanwair2_Amt" +
                $"Group_Amt,Lanwair4_Amt,Retik_Amt," +
                $"ConcessionTicket,ConcessionPoint," +
                $"SOneConcessionReBack,YouBickToLrt," +
                $"AllPassPayback,RideSubsidy";
        }

        public override string ToString()
        {
            return $"{Electront_Amt}," +
                $"{Student_Amt}," +
                $"{Welfare_Amt}," +
                $"{Credit_Amt}," +
                $"{Qrcode_Amt}," +
                $"{SOneTkt_Amt}," +
                $"{SConcessionTkt_Amt}," +
                $"{SBikeTkt_Amt}," +
                $"{OneDay1_Amt}," +
                $"{OneDay_Exit_Amt}," +
                $"{Lanwair1_Amt}," +
                $"{Lanwair2_Amt}," +
                $"{Group_Amt}," +
                $"{Lanwair4_Amt}," +
                $"{Retik_Amt}," +
                $"{ConcessionTicket}," +
                $"{ConcessionPoint}," +
                $"{SOneConcessionReBack}," +
                $"{YouBickToLrt}," +
                $"{AllPassPayback}," +
                $"{RideSubsidy}";
        }

    }

    #endregion

    #region 每日全線各站分時運量日報表

    public class DayEachStationEachTime
    {

    }

    #endregion

    #region 票卡進出站運量

    public class DayElectronicTicket
    {
        public int Station { set; get; }
        public int Electron_Entry_num { set; get; }
        public int Electron_Exit_num { set; get; }
        public int Student_Entry_num { set; get; }
        public int Student_Exit_num { set; get; }
        public int Welfare_Entry_num { set; get; }
        public int Welfare_Exit_num { set; get; }
        public int all_pass_Entry_num { set; get; }
        public int all_pass_Exit_num { set; get; }
        public int all_pass_Student_Entry_num { set; get; }
        public int all_pass_Student_Exit_num { set; get; }
        public int SOneTkt_Entry_num { set; get; }
        public int SOneTkt_Exit_num { set; get; }
        public int SOneTkt_discount_Entry_num { set; get; }
        public int SOneTkt_discount_Exit_num { set; get; }
        public int SOneTkt_bike_Entry_num { set; get; }
        public int SOneTkt_bike_Exit_num { set; get; }
        public int officeman_Entry_num { set; get; }
        public int officeman_Exit_num { set; get; }
        public int OneDay_Entry_num { set; get; }
        public int OneDay_Exit_num { set; get; }
    }   //電票

    #endregion

    #region 營收日報表

    public class DayTrafficAmount
    {

    }

    #endregion

    #region 每日起迄總表

    public class DayOriginDestination
    {
        
    }

    #endregion

    #region 設備營收日報表

    public class DayEquipAmount
    {

    }

    #endregion

    #endregion

}
