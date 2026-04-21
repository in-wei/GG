using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//--
using System.Windows.Forms;
using UsuallyCommond.MyEnum;

namespace GlobalCommond.ViewModel
{

    public class ThreadList
    {
        public string Name { get; set; }
        public System.Threading.Thread thread { get; set; }
        public bool IsRun { get; set; }

        public ThreadList()
        {
            Name = "";
            //thread = new System.Threading.Thread(() => {});
            IsRun = false;
        }

        public override string ToString()
        {
            return $"Name:{Name}, IsRun:{IsRun}";
        }

        

    }

    public class TicketType
    {
        public string TicketName { get; set; }
        public string Check { get; set; }
        public string Volume { get; set; }
        public string Amount { get; set; }
        public string Code { get; set; }

        public TicketType()
        {
            TicketName = "";
            Check = "";
            Volume = "";
            Amount = "";
            Code = "";
        }

        public override string ToString()
        {
            return $"{TicketName},{Check},{Volume},{Amount},{Code}";
        }

        public string ExportTitle_English()
        {
            return "TicketName,Check,Volume,Amount,Code";
        }
        public string ExportTile_zhTW()
        {
            return "車票總類,檢查,運量,營收,代號";
        }
    }

    public class Month_TxnTransaction
    {
        public string Filename { get; set; }
        public DateTime OperationDate { get; set; }
        public List<TxnTransaction> List_TxnTransaction { get; set; }
        public List<TxnTransactionOd> List_TxnTransactionOd { get; set; }

        public Month_TxnTransaction()
        {
            Filename = "";
            OperationDate = new DateTime();
            List_TxnTransaction = new List<TxnTransaction>();
            List_TxnTransactionOd = new List<TxnTransactionOd>();
        }

        public override string ToString()
        {
            return $"FileName:{Filename}, OperationDate:{OperationDate}, List_TxnTransaction.Count:{List_TxnTransaction.Count}, List_TxnTransactionOd.Count:{List_TxnTransactionOd.Count}";
        }
    }

    public class ReportList
    {
        public LrtTxnClientSwitch methode { get; set; }
        public override string ToString()
        {
            return methode.ToString();
        }
    }

    public class CheckBoxStatus
    {
        public DateTime OperationStart { get; set; }
        public DateTime OperationEnd { get; set; }
        public bool EasyAnalyze { get; set; }
        public bool TxnTypeAnalyze { get; set; }
        public bool IOAnalyze { get; set; }
        public bool CompareAnalyze { get; set; }
        public bool TrnAnalyze { get; set; }
        public bool TPassAnalyze { get; set; }
        public bool VolumnReportAnalyze { get; set; }
        public bool AddValueAnalyze { get; set; }
        public bool EveryIssuerAnalyze { get; set; }
        public bool checkbox10 { get; set; }
        public bool checkbox11 { get; set; }
        public bool checkbox12 { get; set; }
        public bool checkbox13 { get; set; }
        public bool checkbox14 { get; set; }
        public bool checkbox15 { get; set; }
        public bool checkbox16 { get; set; }
        public bool checkbox17 { get; set; }
        public bool checkbox18 { get; set; }
        public bool checkbox19 { get; set; }
        public bool checkbox20 { get; set; }
    }

    public class CheckBoxList
    {
        public LrtTxnClientSwitch Methode { get; set; }
        public string MethodeString { get; private set; }
        public CheckBox CB { get; set; }

        public CheckBoxList()
        {
            //while(MethodeString == null)
            //{
            //    try
            //    {
            //        MethodeString = Methode.ToDescription();
            //        CB.Text = MethodeString;
            //    }
            //    catch (Exception ex)
            //    {
            //        Console.WriteLine(ex.Message);
            //    }
            //}
        }
    }

    public class ReportFileTemp2
    {
        public string TransportSystem { get; set; }
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
            return $"系統,路線," +
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
            return $"TransportSystem,OperationLine," +
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
            return $"{TransportSystem},{OperationLine}," +
                $"{ReportNameCh},{ReportNameEng}," +
                $"{StartColumn},{ReportDateRow}," +
                $"{StartColumn},{StartRow}," +
                $"{EndColumn},{EndRow}," +
                $"{BreakOffRow},{BreakOffRowNum}," +
                $"{BreakOffColumn},{BreakOffColumnNum}," +
                $"{MobilePayColumn},{OwnTiketColumn},";
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
        public string SAM_ID { get; set; }

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
            $"{LAST_UTILISATION_DATE}," +
            $"{SAM_ID},";

        public string ExportTitle_English() => $"" +
            $"CARD_TXN_TYPE_ID," +
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
            $"LAST_UTILISATION_DATE," +
            $"SAM_ID,";

        public string ExportTitle_zhTW() => $"" +
            $"交易類別," +
            $"次交易類別," +
            $"設備編號," +
            $"交易時間," +
            $"卡號," +
            $"業者別," +
            $"交易序號," +
            $"交易金額," +
            $"卡片餘額," +
            $"交易車站," +
            $"清分日," +
            $"營運日," +
            $"進站車站," +
            $"轉乘優惠代碼," +
            $"轉乘優惠金額(GATE使用)," +
            $"身分優惠金額," +
            $"罰款," +
            $"社福使用點數/常客優惠累積次數," +
            $"社福剩餘點數/常客優惠累積金額," +
            $"悠遊卡特種票票種代碼," +
            $"進站時間," +
            $"區域碼," +
            $"自有票種票種代碼," +
            $"轉乘優惠金額(PAM使用)," +
            $"交易身分(PAM使用)," +
            $"定期票續購起始日," +
            $"定期票購票日(首用日)," +
            $"定期票到期日," +
            $"判別綁定金融業者的虛擬卡號,";

        #endregion
    } // 捷運交易資料原始載入

    public class EachSellTitle
    {
        public string S1 { get; set; }
        public string S2 { get; set; }

        public override string ToString() => $"{S1},{S2}";

        public string ExportTitle_zhTW() => "程式欄位,班結表欄位";

    }

    public class FileLastName
    {
        public string LastName { get; set; }
        public string Mark { get; set; }
        public string Suggestions { get; set; }

        public override string ToString() => $"{LastName},{Mark},{Suggestions}";

        public string ExportTitle() => "附檔名,描述,建議開啟軟體";
    }

    public class ListTxn
    {
        public DateTime OperationDate { get; set; }
        public List<TxnTransaction> txnTransaction { get; set; }
    }

    public class LoadData
    {
        public string No { get; set; }
        public string FileName { get; set; }
        public string InstallDate { get; set; }
        public string StartDate { get; set; }
        public string FinishDate { get; set; }

        public LoadData()
        {
            No = "";
            FileName = "";
            StartDate = "";
            FinishDate = "";
        }

        #region Methods

        public override string ToString() => $"" +
            $"{No}," +
            $"{FileName}," +
            $"{InstallDate}," +
            $"{StartDate}," +
            $"{FinishDate},";

        public string ExportTitle_English() => $"" +
            $"No," +
            $"FileName," +
            $"InstallDate," +
            $"StartDate," +
            $"FinishDate,";

        public string ExportTitle_zhTW() => $"" +
            $"號," +
            $"檔案名," +
            $"載入日期," +
            $"載入時間," +
            $"完成時間,";

        #endregion

    }

    public class TxnTransactionOd
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
        public string SAM_ID { get; set; }
        public string M_Entry { get; set; }
        public string M_Exit { get; set; }
        public string TicketType { get; set; }
        public string TransferMark { get; set; }
        public string Remark { get; set; }
        public string AmountRemark { get; set; }
        public string ShoudPay { get; set; }
        public string RealPay { get; set; }
        public string AllPassAmount { get; set; }
        public string DiscAmount { get; set; }
        public string DiscPoint { get; set; }
        public string TrnNormalAmount { get; set; }
        public string TrnAllPassAmount { get; set; }
        public string Off_10 { get; set; }
        public string DistanceRatio { get; set; }
        public string Split_ShoudPay { get; set; }
        public string Split_RealPay { get; set; }
        public string Split_AllPassAmount { get; set; }
        public string Split_DiscAmount { get; set; }
        public string Split_DiscPoint { get; set; }
        public string Split_TrnNormalAmount { get; set; }
        public string Split_TrnAllPassAmount { get; set; }
        public string Split_Off_10 { get; set; }
        public double TicketSubsidy { get; set; } // 補貼

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
            $"{LAST_UTILISATION_DATE}," +
            $"{SAM_ID}," +
            $"{M_Entry}," +
            $"{M_Exit}," +
            $"{TicketType}," +
            $"{TransferMark}," +
            $"{Remark}," +
            $"{AmountRemark}," +
            $"{ShoudPay}," +
            $"{RealPay}," +
            $"{AllPassAmount}," +
            $"{DiscAmount}," +
            $"{DiscPoint}," +
            $"{TrnNormalAmount}," +
            $"{TrnAllPassAmount}," +
            $"{Off_10}," +
            $"{DistanceRatio}," +
            $"{Split_ShoudPay}," +
            $"{Split_RealPay}," +
            $"{Split_AllPassAmount}," +
            $"{Split_DiscAmount}," +
            $"{Split_DiscPoint}," +
            $"{Split_TrnNormalAmount}," +
            $"{Split_TrnAllPassAmount}," +
            $"{Split_Off_10}," +
            $"{TicketSubsidy},";

        public string ExportTitle_English() => $"" +
            $"CARD_TXN_TYPE_ID," +
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
            $"LAST_UTILISATION_DATE," +
            $"SAM_ID," +
            $"M_Entry," +
            $"M_Exit," +
            $"TicketType," +
            $"TransferMark," +
            $"Remark," +
            $"AmountRemark," +
            $"ShoudPay," +
            $"RealPay," +
            $"AllPassAmount," +
            $"DiscAmount," +
            $"DiscPoint," +
            $"TrnNormalAmount," +
            $"TrnAllPassAmount," +
            $"Off_10," +
            $"DistanceRatio," +
            $"Split_ShoudPay," +
            $"Split_RealPay," +
            $"Split_AllPassAmount," +
            $"Split_DiscAmount," +
            $"Split_DiscPoint," +
            $"Split_TrnNormalAmount," +
            $"Split_TrnAllPassAmount," +
            $"Split_Off_10," +
            $"TicketSubsidy,";

        public string ExportTitle_zhTW() => $"" +
            $"交易類別," +
            $"次交易類別," +
            $"設備編號," +
            $"交易時間," +
            $"卡號," +
            $"業者別," +
            $"交易序號," +
            $"交易金額," +
            $"卡片餘額," +
            $"交易車站," +
            $"清分日," +
            $"營運日," +
            $"進站車站," +
            $"轉乘優惠代碼," +
            $"轉乘優惠金額(GATE使用)," +
            $"身分優惠金額," +
            $"罰款," +
            $"社福使用點數/常客優惠累積次數," +
            $"社福剩餘點數/常客優惠累積金額," +
            $"悠遊卡特種票票種代碼," +
            $"進站時間," +
            $"區域碼," +
            $"自有票種票種代碼," +
            $"轉乘優惠金額(PAM使用)," +
            $"交易身分(PAM使用)," +
            $"定期票續購起始日," +
            $"定期票購票日(首用日)," +
            $"定期票到期日," +
            $"判別綁定金融業者的虛擬卡號," +
            $"環狀線進站," +
            $"環狀線出站," +
            $"車票種類," +
            $"轉乘標記," +
            $"備註," +
            $"營收標記," +
            $"應收金額," +
            $"實收金額," +
            $"公共運輸定期票票收," +
            $"社福優惠," +
            $"社福點數," +
            $"轉乘優惠," +
            $"定期票轉乘," +
            $"社福吸收," +
            $"距離比例," +
            $"應收金額(拆帳)," +
            $"實收金額(拆帳)," +
            $"公共運輸定期票票收(拆帳)," +
            $"社福優惠(拆帳)," +
            $"社福點數(拆帳)," +
            $"轉乘優惠(拆帳)," +
            $"定期票轉乘(拆帳)," +
            $"社福吸收(拆帳)," +
            $"票差補貼,";

        #endregion
    } // 捷運交易資料修改

    public class CardId
    {
        #region Properties

        public string CardSN { get; set; }
        public int Count { get; set; }

        #endregion

        #region Methods

        public override string ToString() => $"{CardSN},{Count}";

        public string ExportTitle_English() => "CardSN,Count";

        #endregion
    }
    /**
    public class MrtStationList_Old
    {
        public string CodeName { get; set; }
        public string StationName { get; set; }
        public string CodeName2 { get; set; }
        public string StationName2 { get; set; }

        public override string ToString() => $"{CodeName},{StationName},{CodeName2},{StationName2}";

        public string ExportTitle_English() => "CodeName,StationName,CodeName2,StationName2";

        public string ExportTitle_zhTW() => "代號,車站,代號2,車站2";

    }
    **/
    public class MrtStationList
    {
        public string Company1 { get; set; }
        public string CodeName1 { get; set; }
        public string StationName1 { get; set; }
        public string Company2 { get; set; }
        public string CodeName2 { get; set; }
        public string StationName2 { get; set; }

        public override string ToString() => $"{Company1},{CodeName1},{StationName1},{Company2},{CodeName2},{StationName2},";

        public string ExportTitle_English() => "Compaany1,CodeName1,StationName1,Company2,CodeName2,StationName2,";

        public string ExportTitle_zhTW() => "公司1,代號1,車站1,公司2,代號2,車站2,";

    }
    /**
    public class MrtStationOd
    {
        public string EntryCompany { get; set; }
        public string EntryStation { get; set; }
        public string ExitCompany { get; set; }
        public string ExitStation { get; set; }
        public string IsCount { get; set; }
        public string Divide { get; set; }
        public string Amt { get; set; }

        public override string ToString() => $"" +
                $"{EntryCompany}," +
                $"{EntryStation}," +
                $"{ExitCompany}," +
                $"{ExitStation}," +
                $"{IsCount}," +
                $"{Divide}," +
                $"{Amt},";

        public string ExportTitle_English() => $"" +
                $"EntryCompany," +
                $"EntryStation," +
                $"ExitCompany," +
                $"ExitStation," +
                $"IsCount," +
                $"Divide," +
                $"Amt,";

        public string ExportTitle_zhTW() => $"" +
                $"起站系統," +
                $"起站," +
                $"訖站系統," +
                $"訖站," +
                $"計算," +
                $"比例," +
                $"運價,";

    }
    **/
    public class MrtStationOdList
    {
        public string Operation_Area { get; set; }
        public string operation_Line { get; set; }
        public string Entry { get; set; }
        public string Exit { get; set; }
        public string Amount_Full { get; set; }
        public string Odamt { get; set; }
        public string Amount_Walfare { get; set; }
        public string Amount_Child { get; set; }
        public string MileAge_Full { get; set; }
        public string IsPass { get; set; }
        public string MileAge_Own { get; set; }
        public string MileAge_Scale { get; set; }
        public string Entry_Own { get; set; }
        public string Exit_Own { get; set; }

        public override string ToString() => $"{Operation_Area}," +
                $"{operation_Line}," +
                $"{Entry}," +
                $"{Exit}," +
                $"{Amount_Full}," +
                $"{Odamt}," +
                $"{Amount_Walfare}," +
                $"{Amount_Child}," +
                $"{MileAge_Full}," +
                $"{IsPass}," +
                $"{MileAge_Own}," +
                $"{MileAge_Scale}," +
                $"{Entry_Own}," +
                $"{Exit_Own},";

        public string ExportTitle_English() => "";

        public string ExportTitle_zhTW() => $"營運路線," +
                $"路線代號," +
                $"起站," +
                $"訖站," +
                $"全票票價," +
                $"票差," +
                $"敬老卡愛心卡愛心陪伴卡及新北市兒童優惠票價," +
                $"臺北市兒童優惠票價," +
                $"距離(公里)," +
                $"是否經過環狀線," +
                $"環狀線距離(公里)," +
                $"環狀線距離比例," +
                $"環狀線起站," +
                $"環狀線訖站,";

    }

    public class String_String_String_Int
    {
        public string ChName { get; set; }
        public string EngName { get; set; }
        public string VolumnCheck { get; set; }
        public int Num { get; set; }

        public override string ToString() => $"" +
                $"{ChName}," +
                $"{EngName}," +
                $"{VolumnCheck}," +
                $"{Num}";

        public string ExportTitle_English() => $"" +
                $"ChName," +
                $"EngName," +
                $"VolumnCheck," +
                $"Num";

        public string ExportTitle_zhTW() => $"" +
                $"中文名稱," +
                $"英文名稱," +
                $"運量計算," +
                $"數值";

    }

    public class String_String_Int_Int_Int
    {
        public string ChName { get; set; }
        public string EngName { get; set; }
        public int IssuerRawTag { get; set; }
        public int AddValueTag { get; set; }
        public int Num { get; set; }

        public override string ToString() => $"" +
                $"{ChName}," +
                $"{EngName}," +
                $"{IssuerRawTag}," +
                $"{AddValueTag}," +
                $"{Num}";

        public string ExportTitle_English() => $"" +
                $"ChName," +
                $"EngName," +
                $"IssuerRawTag," +
                $"AddValueTag," +
                $"Num";

        public string ExportTitle_zhTW() => $"" +
                $"中文名稱," +
                $"英文名稱," +
                $"業者運量標記," +
                $"加值標記," +
                $"數值";

    }

    public class String_String_String_String_Int
    {
        public string ChName { get; set; }
        public string EngName { get; set; }
        public string mCheck { get; set; }
        public string VolumnCheck { get; set; }
        public int Num { get; set; }

        public override string ToString() => $"" +
                $"{ChName}," +
                $"{EngName}," +
                $"{mCheck}," +
                $"{VolumnCheck}," +
                $"{Num},";

        public string ExportTitle_English() => $"" +
                $"ChName," +
                $"EngName," +
                $"Check," +
                $"VolumnCheck," +
                $"Num,";

        public string ExportTitle_zhTW() => $"" +
                $"中文名稱," +
                $"英文名稱," +
                $"檢測," +
                $"運量計算," +
                $"數值,";

    }

    public class OnlyInt
    {
        public int Num { get; set; }

        public override string ToString() => $"{Num},";

        public string ExportTitle_English() => "Num,";

        public string ExportTitle_zhTW() => $"數值";

    }

    public class TxnTransactionMegre
    {
        public string BUSINESS_DATE { get; set; }
        public string SVCE_LOC_ID { get; set; }
        public string EQUIP_TYPE { get; set; }
        public string EQUIP_ID { get; set; }
        public string CARD_TXN_TYPE_ID { get; set; }
        public string CARD_TXN_SUBTYPE_ID { get; set; }
        public string ISSUER_ID { get; set; }
        public string FARE_PRODUCT_TYPE_ID { get; set; }
        public string TICKET_SUBTYPE_ID { get; set; }
        public int Count { get; set; }
        public int Amt { get; set; }

        public override string ToString() => $"" +
                $"{BUSINESS_DATE:yyyy/MM/dd}," +
                $"{SVCE_LOC_ID}," +
                $"{EQUIP_TYPE}," +
                $"{EQUIP_ID}," +
                $"{ISSUER_ID}," +
                $"{CARD_TXN_TYPE_ID}," +
                $"{CARD_TXN_SUBTYPE_ID}," +
                $"{TICKET_SUBTYPE_ID}, " +
                $"{FARE_PRODUCT_TYPE_ID}," +
                $"{Count}," +
                $"{Amt}" +
                $"";

        public string ExportTitle_English() => $"" +
                $"BUSINESS_DATE," +
                $"LOCATION_ID," +
                $"EQUIP_TYPE," +
                $"EQUIP_ID," +
                $"ISSUER_ID," +
                $"TXN_TYPE," +
                $"TXN_SUBTYPE," +
                $"TICKET_SUBTYPE, " +
                $"FARE_PRODUCT_TYPE," +
                $"Count," +
                $"Amt" +
                $"";

        public string ExportTitle_zhTW() => $"" +
                $"營運日," +
                $"車站," +
                $"設備種類," +
                $"設備編號," +
                $"業者別," +
                $"交易類別," +
                $"交易子類別," +
                $"自有票特種票代碼, " +
                $"電子票特種票代碼," +
                $"數量," +
                $"金額" +
                $"";

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

        public override string ToString() => $"" +
                $"{ReportNameCh}," +
                $"{ReportNameEng}," +
                $"{ReportDateColumn}," +
                $"{ReportDateRow}," +
                $"{StartColumn}," +
                $"{StartRow}," +
                $"{EndColumn}," +
                $"{EndRow}," +
                $"{BreakOffRow}," +
                $"{BreakOffRowNum}," +
                $"{BreakOffColumn}," +
                $"{BreakOffColumnNum},";

        public string ExportTitle_English() => $"" +
                $"ReportNameCh," +
                $"ReportNameEng," +
                $"ReportDateColumn," +
                $"ReportDateRow," +
                $"StartColumn," +
                $"StartRow," +
                $"EndColumn," +
                $"EndRow," +
                $"BreakOffRow," +
                $"BreakOffRowNum," +
                $"BreakOffColumn," +
                $"BreakOffColumnNum,";

        public string ExportTitle_zhTW() => $"" +
                $"資料表中文," +
                $"資料表英文," +
                $"產製日期欄(英)," +
                $"產製日期行(數)," +
                $"資料起始欄(英)," +
                $"資料起始行(數)," +
                $"資料結束欄(英)," +
                $"資料結束行(數)," +
                $"資料中斷行(數)," +
                $"中斷行數," +
                $"資料中斷欄(英)," +
                $"中斷欄數,";

    }

    public class OdSubsidy
    {
        public string EntryStation { get; set; }
        public string ExitStation { get; set; }
        public int Num { get; set; }

        public override string ToString() => $"" +
                $"{EntryStation}," +
                $"{ExitStation}," +
                $"{Num},";

        public string ExportTitle_English() => $"" +
                $"EntryStation," +
                $"ExitStation," +
                $"Num,";

        public string ExportTitle_zhTW() => $"" +
                $"進站車站," +
                $"出站車站," +
                $"數量,";

    }

    public class teatime
    {
        public string CARD_TXN_TYPE_ID { get; set; }
        public string CARD_TXN_SUBTYPE_ID { get; set; }
        public string ISSUER_ID { get; set; }
        public string BUSINESS_DATE { get; set; }
        public string ENTRY_LOC_ID { get; set; }
        public string SVCE_LOC_ID { get; set; }
        public string FARE_PRODUCT_TYPE_ID { get; set; }
        public string XFER_CODE { get; set; }
        public int Counts { get; set; }
        public int Amts { get; set; }

        public override string ToString() => $"" +
                $"{CARD_TXN_TYPE_ID}," +
                $"{CARD_TXN_SUBTYPE_ID}," +
                $"{FARE_PRODUCT_TYPE_ID}," +
                $"{ISSUER_ID}," +
                $"{BUSINESS_DATE}," +
                $"{ENTRY_LOC_ID}," +
                $"{SVCE_LOC_ID}," +
                $"{XFER_CODE}," +
                $"{Counts}," +
                $"{Amts},";

        public string ExportTitle_English() => $"" +
                $"CARD_TXN_TYPE_ID," +
                $"CARD_TXN_SUBTYPE_ID," +
                $"FARE_PRODUCT_TYPE_ID," +
                $"ISSUER_ID," +
                $"BUSINESS_DATE," +
                $"ENTRY_LOC_ID," +
                $"SVCE_LOC_ID," +
                $"XFER_CODE," +
                $"Counts," +
                $"Amts,";

        public string ExportTitle_zhTW() => $"" +
                $"交易類別," +
                $"次交易類別," +
                $"悠遊卡特種票票種代碼," +
                $"業者別," +
                $"營運日," +
                $"進站車站," +
                $"交易車站," +
                $"轉乘優惠代碼," +
                $"數量," +
                $"金額,";

    }

    //--
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

        public override string ToString() => $"" +
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

        public string ExportTitle_English() => $"" +
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

        public string ExportTitle_zhTW() => $"" +
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

        public override string ToString() => $"{Operation}," +
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

        public string ExportTitle_zhTW() => $"營運," +
                $"車站," +
                $"" +
                $"一日票," +
                $"優待票10元,優待票12元,優待票15元," +
                $"單程票20元,單程票25元,單程票30元," +
                $"自行車票,通車紀念套票,行政院通勤月票紀念卡," +
                $"團體票張數,團體票人數,團體票金額," +
                $"活動1,活動2," +
                $"設備取出,50倍車資," +
                $"" +
                $"兌換單程票20元,兌換單程票25元,兌換單程票30元," +
                $"一日票(KKDay),一日票(新北幣),切票,兌換一日票," +
                $"" +
                $"優待票代墊款,退費通知,自動售票機退費," +
                $"刷卡機退費,信用卡退費,乘車碼退費," +
                $"退卡,團體票退費";

        public string ExportTitle_English() => $"OperationDate," +
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

    public class LrtStationList
    {
        public string OperationLine { get; set; }
        public string CodeNumber { get; set; }
        public string CodeName { get; set; }
        public string StationName { get; set; }
        public string StationPao { get; set; }

        public string ExportTitle_zhTW() => $"路線,代碼,代號,名稱,詢問處";

        public string ExportTitle_English() => $"OperationLine,CodeNumber,CodeName,StationName,StationPao";

        public override string ToString() => $"{OperationLine},{CodeNumber},{CodeName},{StationName},{StationPao}";

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

        public string ExportTitle_zhTW() => $"" +
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

        public string ExportTitle_English() => $"" +
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

        public override string ToString() => $"" +
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

    public class Subsidy
    {
        public string Operation_Area { get; set; }
        public string Operation_Line { get; set; }
        public string StartStation { get; set; }
        public string EndStation { get; set; }
        public string ShoudPay { get; set; }
        public string OdAmt { get; set; }

        public string ExportTitle_English() => $"" +
                $"Operation_Area," +
                $"Operation_Line," +
                $"StartStation," +
                $"EndStation," +
                $"ShoudPay," +
                $"OdAmt";

        public string ExportTitle_zhTW() => $"" +
                $"營運區域," +
                $"路線代號," +
                $"起站," +
                $"訖站," +
                $"運價," +
                $"差額";

        public override string ToString() => $"" +
                $"{Operation_Area}," +
                $"{Operation_Line}," +
                $"{StartStation}," +
                $"{EndStation}," +
                $"{ShoudPay}," +
                $"{OdAmt}";

    }

    public class ReSql
    {
        public string Using { get; set; }
        public bool UpData { get; set; }

        public string ExportTitle_zhTW() => $"使用語法,是否更新";

        public string ExportTitle_English() => $"Using,UpData";

        public override string ToString() => $"{Using},{UpData}";

    }

    public class SqlList
    {
        public string UseNameCh { get; set; }
        public string FileName1 { get; set; }
        public string FileName2 { get; set; }
        public string Remark { get; set; }
        public string Using { get; set; }

        public string ExportTitle_zhTW() => $"報表中文,語法名稱1,語法名稱2,停用標記,使用語法";

        public string ExportTitle_English() => $"UseNameCh,FileName1,FileName2,Remark,Using";

        public override string ToString() => $"{UseNameCh},{FileName1},{FileName2},{Remark},{Using}";

    }

    public class MobileCardType
    {
        public string CardName { get; set; }
        public string CardType { get; set; }
        public string CardSubType { get; set; }

        public string ExportTitle_zhTW() => $"卡片名稱,卡種,卡別";

        public string ExportTitle_English() => $"CardName,CardType,CardSubType";

        public override string ToString() => $"{CardName},{CardType},{CardSubType}";

    }

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

        public string ExportTitle_zhTW() => $"卡號," +
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

        public string ExportTitle_English() => $"" +
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

        public override string ToString() => $"" +
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

    public class MobilePayVolum
    {
        public const string Normal      = "票價扣款";
        public const string NoEnrty     = "進站補扣款";            //出站時產生  (沒刷進站刷出站)
        public const string NoExit      = "出站補扣款";            //半夜結帳產生(有刷進站沒出站)
        public const string OutTime     = "逾時罰款";              //進站後很久才出站
        public const string InTime      = "同站進出時間內不扣款";  //進站後在很短的時間內刷出站
        public const string NextDayExit = "跨日出站扣款";          //同正常進出，只是跨日了
    }

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
                case "普通卡": ColumnValue = 0; break;
                case "學生卡": ColumnValue = 1; break;
                case "優待卡": ColumnValue = 2; break;
                case "一般交通月票": ColumnValue = 3; break;
                case "學生交通月票": ColumnValue = 4; break;
                case "信用卡": ColumnValue = 5; break;
                case "乘車碼": ColumnValue = 6; break;
                case "一般單程票": ColumnValue = 7; break;
                case "優待單程票": ColumnValue = 8; break;
                case "自行車單程票": ColumnValue = 9; break;
                case "一日票": ColumnValue = 10; break;
                case "一日電票": ColumnValue = 11; break;
                case "活動1": ColumnValue = 12; break;
                case "活動2": ColumnValue = 13; break;
                case "團體票張數": ColumnValue = 14; break;
                case "團體票人數": ColumnValue = 15; break;
                case "員工卡": ColumnValue = 16; break;
                default: ColumnValue = -1; break;
            }
            return ColumnValue;
        }

        public int ColumnCompare_English(string GetTitle)
        {
            int ColumnValue;
            switch (GetTitle)
            {
                case "Electron_Exit_num": ColumnValue = 0; break;
                case "Student_Exit_num": ColumnValue = 1; break;
                case "Welfare_Exit_num": ColumnValue = 2; break;
                case "All_Pass_Common_Exit_num": ColumnValue = 3; break;
                case "All_Pass_Student_Exit_num": ColumnValue = 4; break;
                case "Credit_num": ColumnValue = 5; break;
                case "Qrcode_num": ColumnValue = 6; break;
                case "SOneTkt_num": ColumnValue = 7; break;
                case "SConcessionTkt_num": ColumnValue = 8; break;
                case "SBikeTkt_num": ColumnValue = 9; break;
                case "OneDay1_num": ColumnValue = 10; break;
                case "OneDay_Exit_num": ColumnValue = 11; break;
                case "Lanwair1_num": ColumnValue = 12; break;
                case "Lanwair2_num": ColumnValue = 13; break;
                case "Group_Tik_num": ColumnValue = 14; break;
                case "Group_Peo_num": ColumnValue = 15; break;
                case "Office_Staff_num": ColumnValue = 16; break;
                default: ColumnValue = -1; break;
            }
            return ColumnValue;
        }

        public string ExportTitle_zhTW() => $"普通卡," +
            $"學生卡," +
            $"優待卡," +
            $"一般交通月票," +
            $"學生交通月票," +
            $"信用卡," +
            $"乘車碼," +
            $"一般單程票," +
            $"優待單程票," +
            $"自行車單程票," +
            $"一日票," +
            $"一日電票," +
            $"活動1," +
            $"活動2," +
            $"團體票張數," +
            $"團體票人數," +
            $"員工卡";

        public string ExportTitle_English() => $"Electron_Exit_num," +
            $"Student_Exit_num," +
            $"Welfare_Exit_num," +
            $"All_Pass_Common_Exit_num," +
            $"All_Pass_Student_Exit_num," +
            $"Credit_num," +
            $"Qrcode_num," +
            $"SOneTkt_num," +
            $"SConcessionTkt_num," +
            $"SBikeTkt_num," +
            $"OneDay1_num," +
            $"OneDay_Exit_num" +
            $"Lanwair1_num," +
            $"Lanwair2_num," +
            $"Group_Tik_num," +
            $"Group_Peo_num," +
            $"Office_Staff_num";

        public override string ToString() => $"{Electron_Exit_num}," +
            $"{Student_Exit_num}," +
            $"{Welfare_Exit_num}," +
            $"{All_Pass_Common_Exit_num}," +
            $"{All_Pass_Student_Exit_num}," +
            $"{Credit_num}," +
            $"{Qrcode_num}," +
            $"{SOneTkt_num}," +
            $"{SConcessionTkt_num}," +
            $"{SBikeTkt_num}," +
            $"{OneDay1_num}," +
            $"{OneDay_Exit_num}," +
            $"{Lanwair1_num}," +
            $"{Lanwair2_num}," +
            $"{Group_Tik_num}," +
            $"{Group_Peo_num}," +
            $"{Office_Staff_num}";

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
                case "普通票": ColumnValue = 0; break;
                case "學生票": ColumnValue = 1; break;
                case "優待票": ColumnValue = 2; break;
                case "信用卡": ColumnValue = 3; break;
                case "乘車碼": ColumnValue = 4; break;
                case "一般單程票": ColumnValue = 5; break;
                case "優待單程票": ColumnValue = 6; break;
                case "自行車單程票": ColumnValue = 7; break;
                case "一日票": ColumnValue = 8; break;
                case "一日電票": ColumnValue = 9; break;
                case "活動1": ColumnValue = 10; break;
                case "活動2": ColumnValue = 11; break;
                case "團體票": ColumnValue = 12; break;
                case "活動4": ColumnValue = 13; break;
                case "補票": ColumnValue = 14; break;
                case "社福票": ColumnValue = 15; break;
                case "社福點數": ColumnValue = 16; break;
                case "社福代墊款": ColumnValue = 17; break;
                case "微笑單車轉乘": ColumnValue = 18; break;
                case "交通月票": ColumnValue = 19; break;
                case "里程補貼": ColumnValue = 20; break;
                default: ColumnValue = -1; break;
            }
            return ColumnValue;
        }

        public int ColumnCompare_English(string GetTitle)
        {
            int ColumnValue;
            switch (GetTitle)
            {
                case "Electront_Amt": ColumnValue = 0; break;
                case "Student_Amt": ColumnValue = 1; break;
                case "Welfare_Amt": ColumnValue = 2; break;
                case "Credit_Amt": ColumnValue = 3; break;
                case "Qrcode_Amt": ColumnValue = 4; break;
                case "SOneTkt_Amt": ColumnValue = 5; break;
                case "SConcessionTkt_Amt": ColumnValue = 6; break;
                case "SBikeTkt_Amt": ColumnValue = 7; break;
                case "OneDay1_Amt": ColumnValue = 8; break;
                case "OneDay_Exit_Amt": ColumnValue = 9; break;
                case "Lanwair1_Amt": ColumnValue = 10; break;
                case "Lanwair2_Amt": ColumnValue = 11; break;
                case "Group_Amt": ColumnValue = 12; break;
                case "Lanwair4_Amt": ColumnValue = 13; break;
                case "Retik_Amt": ColumnValue = 14; break;
                case "ConcessionTicket": ColumnValue = 15; break;
                case "ConcessionPoint": ColumnValue = 16; break;
                case "SOneConcessionReBack": ColumnValue = 17; break;
                case "YouBickToLrt": ColumnValue = 18; break;
                case "AllPassPayback": ColumnValue = 19; break;
                case "RideSubsidy": ColumnValue = 20; break;
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

    #endregion

}
