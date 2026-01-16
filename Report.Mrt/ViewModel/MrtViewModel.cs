using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//--
using UsuallyCommond.MyEnum;

namespace Report.Mrt.ViewModel
{
    public class MrtViewModel
    {

    }

    public class PublicUse
    {
        public string Receive { get; set; }
        public ExecutionMode ExportMode { get; set; }
        public string Path { get; set; }
        public string Coung { get; set; }
        public string Msg { get; set; }
        
    }

    public class ReportList
    {
        public LrtTxnClientSwitch methode { get; set; }
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
        public bool checkbox7 { get; set; }
        public bool AccurateAnalyze { get; set; }
        public bool checkbox9 { get; set; }
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
            $"{LAST_UTILISATION_DATE},";

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
            $"LAST_UTILISATION_DATE,";

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
            $"定期票到期日,";

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
        public string M_Entry { get; set; }
        public string M_Exit { get; set; }
        public string TransferMark { get; set; }
        public string Remark { get; set; }

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
            $"{M_Entry}," +
            $"{M_Exit}," +
            $"{TransferMark}," +
            $"{Remark},";

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
            $"M_Entry," +
            $"M_Exit," +
            $"TransferMark," +
            $"Remark,";

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
            $"環狀線進站," +
            $"環狀線出站," +
            $"轉乘標記," +
            $"備註,";

        #endregion
    }

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

    public class ExportConfig
    {
        #region Properties

        public string Path1 { get; set; }
        public DateTime DTime { get; set; }
        public string FirstName { get; set; }
        public string FileName { get; set; }
        public string SubName { get; set; }
        public string LastName { get; set; }

        #endregion

        #region Methods

        public override string ToString() => $"{Path1},{DTime:yyyy/MM/dd},{FirstName},{FileName},{SubName},{LastName}";

        public string ExportTitle_English() => "Path1,DTime,FirstName,FileName,SubName,LastName";

        #endregion
    }

    public class MrtStationList_Old
    {
        public string CodeName { get; set; }
        public string StationName { get; set; }
        public string CodeName2 { get; set; }
        public string StationName2 { get; set; }

        public override string ToString()
        {
            return $"{CodeName},{StationName},{CodeName2},{StationName2}";
        }

        public string ExportTitle_English()
        {

            return "CodeName,StationName,CodeName2,StationName2";
        }

        public string ExportTitle_zhTW()
        {

            return "代號,車站,代號2,車站2";
        }
        ///   ExportTitle = $"CodeName,StationName,CodeName2,StationName2";

    }

    public class MrtStationList
    {
        public string CodeName { get; set; }
        public string StationName { get; set; }
        public string Company { get; set; }

        public override string ToString()
        {
            return $"{CodeName},{StationName},{Company},";
        }

        public string ExportTitle_English()
        {

            return "CodeName,StationName,Compaany,";
        }

        public string ExportTitle_zhTW()
        {

            return "代號,車站,公司";
        }

    }

    public class MrtStationOd
    {
        public string EntryCompany { get; set; }
        public string EntryStation { get; set; }
        public string ExitCompany { get; set; }
        public string ExitStation { get; set; }
        public string IsCount { get; set; }
        public string Divide { get; set; }
        public string Amt { get; set; }

        public override string ToString()
        {
            return $"" +
                $"{EntryCompany}," +
                $"{EntryStation}," +
                $"{ExitCompany}," +
                $"{ExitStation}," +
                $"{IsCount}," +
                $"{Divide}," +
                $"{Amt},";
        }

        public string ExportTitle_English()
        {

            return $"" +
                $"EntryCompany," +
                $"EntryStation," +
                $"ExitCompany," +
                $"ExitStation," +
                $"IsCount," +
                $"Divide," +
                $"Amt,";
        }

        public string ExportTitle_zhTW()
        {

            return $"" +
                $"起站系統," +
                $"起站," +
                $"訖站系統," +
                $"訖站," +
                $"計算," +
                $"比例," +
                $"運價,";
        }


    }

    public class MrtStationOdList
    {
        public string Entry { get; set; }
        public string Exit { get; set; }
        public string Amount_Full { get; set; }
        public string Amount_Walfare { get; set; }
        public string Amount_Child { get; set; }
        public string MileAge_Full { get; set; }
        public string IsPass { get; set; }
        public string MileAge_Own { get; set; }
        public string MileAge_Scale { get; set; }
        public string Entry_Own { get; set; }
        public string Exit_Own { get; set; }

        public override string ToString()
        {
            return $"" +
                $"{Entry}," +
                $"{Exit}," +
                $"{Amount_Full}," +
                $"{Amount_Walfare}," +
                $"{Amount_Child}," +
                $"{MileAge_Full}," +
                $"{IsPass}," +
                $"{MileAge_Own}," +
                $"{MileAge_Scale}," +
                $"{Entry_Own}," +
                $"{Exit_Own},";
        }

        public string ExportTitle_English()
        {
            return "";
        }

        public string ExportTitle_zhTW()
        {
            return $"" +
                $"起站," +
                $"訖站," +
                $"全票票價," +
                $"敬老卡愛心卡愛心陪伴卡及新北市兒童優惠票價," +
                $"臺北市兒童優惠票價," +
                $"距離(公里)," +
                $"是否經過環狀線," +
                $"環狀線距離(公里)," +
                $"環狀線距離比例," +
                $"環狀線起站," +
                $"環狀線訖站,";
        }

    }

    public class String_String_Int
    {
        public string ChName { get; set; }
        public string EngName { get; set; }
        public int Num { get; set; }

        public override string ToString()
        {
            return $"" +
                $"{ChName}," +
                $"{EngName}," +
                $"{Num}";
        }

        public string ExportTitle_English()
        {

            return $"" +
                $"ChineseName," +
                $"EnglishName," +
                $"Num";
        }

        public string ExportTitle_zhTW()
        {

            return $"" +
                $"中文名稱," +
                $"英文名稱," +
                $"數值";
        }

    }

    public class String_String_String_Int
    {
        public string ChName { get; set; }
        public string EngName { get; set; }
        public string mCheck { get; set; }
        public int Num { get; set; }

        public override string ToString()
        {
            return $"" +
                $"{ChName}," +
                $"{EngName}," +
                $"{mCheck}," +
                $"{Num},";
        }

        public string ExportTitle_English()
        {

            return $"" +
                $"ChineseName," +
                $"EnglishName," +
                $"Check," +
                $"Num,";
        }

        public string ExportTitle_zhTW()
        {

            return $"" +
                $"中文名稱," +
                $"英文名稱," +
                $"檢測," +
                $"數值";
        }

    }

    public class OnlyInt
    {
        public int Num { get; set; }

        public override string ToString()
        {
            return $"{Num},";
        }

        public string ExportTitle_English()
        {
            return "Num,";
        }

        public string ExportTitle_zhTW()
        {

            return $"" +
                $"數值";
        }

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

        public override string ToString()
        {
            return $"" +
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
        }

        public string ExportTitle_English()
        {

            return $"" +
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
        }

        public string ExportTitle_zhTW()
        {
            return $"" +
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

    }

    public class ReportInformation
    {
        public int RideCountAdd { get; set; }
        public int RideCountMinus { get; set; }
        public int AmtAdd { get; set; }
        public int AmtMinus { get; set; }
        public int Pass1280Ride { get; set; }
        public int Pass1200Ride { get; set; }
        public int Pass1200Set { get; set; }
        public int Pass1200CancelCount { get; set; }
        public int Pass1200CancelAmt { get; set; }

        public override string ToString()
        {
            return $"" +
                $"{RideCountAdd}," +
                $"{RideCountMinus}," +
                $"{AmtAdd}," +
                $"{AmtMinus}," +
                $"{Pass1280Ride}," +
                $"{Pass1200Ride}," +
                $"{Pass1200Set}," +
                $"{Pass1200CancelCount}," +
                $"{Pass1200CancelAmt},";
        }

        public string ExportTitle_English()
        {
            return $"" +
                  $"RideCountAdd," +
                  $"RideCountMinus," +
                  $"AmtAdd," +
                  $"AmtMinus," +
                  $"Pass1280Ride," +
                  $"Pass1200Ride," +
                  $"Pass1200Set," +
                  $"Pass1200CancelCount," +
                  $"Pass1200CancelAmt,";
        }

        public string ExportTitle_zhTW()
        {
            return $"" +
                  $"運量數量(加)," +
                  $"運量數量(減)," +
                  $"金額(加)," +
                  $"金額(減)," +
                  $"公共運輸月票運量," +
                  $"行政院通勤月票運量," +
                  $"行政院通勤月票設定," +
                  $"行政院通勤月票取消數量," +
                  $"行政院通勤月票取消金額,";
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

        public override string ToString()
        {
            return $"" +
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
        }

        public string ExportTitle_English()
        {
            return $"" +
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
        }

        public string ExportTitle_zhTW()
        {
            return $"" +
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

    }

    public class OdSubsidy
    {
        public string EntryStation { get; set; }
        public string ExitStation { get; set; }
        public int Num { get; set; }

        public override string ToString()
        {
            return $"" +
                $"{EntryStation}," +
                $"{ExitStation}," +
                $"{Num},";
        }

        public string ExportTitle_English()
        {

            return $"" +
                $"EntryStation," +
                $"ExitStation," +
                $"Num,";
        }

        public string ExportTitle_zhTW()
        {

            return $"" +
                $"進站車站," +
                $"出站車站," +
                $"數量,";
        }

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

        public override string ToString()
        {
            return $"" +
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
        }

        public string ExportTitle_English()
        {
            return $"" +
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
        }

        public string ExportTitle_zhTW()
        {
            return $"" +
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

    }

}
