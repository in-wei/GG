using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//--
using System.ComponentModel;
using System.Reflection;

namespace UsuallyCommond.MyEnum
{
    public static class EnumExtenstions
    {
        public static string ToDescription(this Enum value)
        {
            return value.GetType()
                .GetRuntimeField(value.ToString())
                .GetCustomAttributes<System.ComponentModel.DescriptionAttribute>()
                .FirstOrDefault()?.Description ?? string.Empty;
        }
    }

    #region 列舉類型

    public enum DateTimeSwitch
    {
        RunDatetime,
        LastDatetime,
    }

    public enum ViewListStatus
    {
        [Description("NULL")] NULL,
        [Description("開始中")] Start,
        [Description("執行中")] Running,
        [Description("未啟動")] Closed,
    }

    /// <summary>
    /// 指示Client端執行程式
    /// </summary>
    public enum MrtTxnClientSwitch
    {
        /* 00 */ [Description("NULL")] NA,
        /* 01 */ [Description("簡易資料(每日)")] EzsyAnalyze,
        /* 02 */ [Description("各交易資料")] TxnTypeAnalyze,
        /* 03 */ [Description("進出站交易")] IOAnalyze,
        /* 04 */ [Description("合併里程")] CompareAnalyze,
        /* 05 */ [Description("轉乘(月報表)")] TrnAnalyze,
        /* 06 */ [Description("通勤月票(月報表)")] TPassAnalyze,
        /* 07 */ [Description("運量報表")] VolumnReportAnalyze,
        /* 08 */ [Description("票卡加值統計表")] Day_AddValue,
        /* 09 */ [Description("各家支付業者運量日報表")] Day_EveryIssuer,
        /* 10 */ [Description("C10")] C10Analyze,
        /* 11 */ [Description("C11")] C11Analyze,
        /* 12 */ [Description("C12")] C12Analyze,
        /* 13 */ [Description("C13")] C13Analyze,
        /* 14 */ [Description("C14")] C14Analyze,
        /* 15 */ [Description("C15")] C15Analyze,
        /* 16 */ [Description("C16")] C16Analyze,
        /* 17 */ [Description("C17")] C17Analyze,
        /* 18 */ [Description("C18")] C18Analyze,
        /* 19 */ [Description("C19")] C19Analyze,
        /* 20 */ [Description("C20")] C20Analyze,
    }

    /// <summary>
    /// DateTime分類
    /// </summary>
    public enum DatePark
    {
        [Description("無")/* */] Null,
        [Description("年")/* */] Year,
        [Description("月")/* */] Month,
        [Description("日")/* */] Day,
        [Description("時")/* */] Hour,
        [Description("分")/* */] Minute,
        [Description("秒")/* */] Second,
    }

    /// <summary>
    /// 指示Client端執行程式
    /// </summary>
    public enum LrtTxnClientSwitch
    {
        [Description("NULL")/*-----------------------*/] NA,
        [Description("營運資料統計表(運量)")/*-------*/] Day_Volume,
        [Description("營運資料統計表(營收)")/*-------*/] Day_Amount,
        [Description("每日全線各站分時運量日報表")/*-*/] Day_EachStation_EachTime,
        [Description("票卡進出站運量日報表")/*-------*/] Day_ElectronicTicket,
        [Description("日運量統計表")/*---------------*/] Day_AllRideList,
        [Description("營收日報表")/*-----------------*/] Day_TrafficAmount,
        [Description("每日起訖站總表")/*-------------*/] Day_OriginDestination,
        [Description("設備營收日報表")/*-------------*/] Day_EquipAmount,
        [Description("月運量統計表")/*---------------*/] Month_AllRideList,
        [Description("每月起訖總表")/*---------------*/] Month_OriginDestination,
        [Description("各家票卡運量月報表(車站)")/*---*/] Month_ElectronicTicket_Station,
        [Description("各家票卡運量月報表(天期)")/*---*/] Month_ElectronicTicket_Day,
        [Description("自有票卡運量月報表(車站)")/*---*/] Month_OwnTicketVolume_Station,
        [Description("自有票卡運量月報表(天期)")/*---*/] Month_OwnTicketVolume_Day,
        [Description("營收月報表(車站)")/*-----------*/] Month_TrafficAmount_Station,
        [Description("營收月報表(天期)")/*-----------*/] Month_TrafficAmount_Day,
        [Description("設備營收月報表(車站)")/*-------*/] Month_EquipAmount_Station,
        [Description("設備營收月報表(天期)")/*-------*/] Month_EquipAmount_Day,
        [Description("營運資料月報(運量)")/*---------*/] Month_Volume,
        [Description("營運資料月報(營收)")/*---------*/] Month_Amount,
    }

    public enum LR_Switch
    {
        [Description("NULL")/* */] NA,
        [Description("左")/*   */] Left,
        [Description("右")/*   */] Right,
    }

    /// <summary>
    /// 結班報表人工項目
    /// </summary>
    public enum LrtPaoServiceList
    {
        [Description("NULL")/*              */] NA,
        [Description("販售_一日票")/*       */] Ticket_OneDay,
        [Description("販售_優待票_10元")/*  */] OneTime10,
        [Description("販售_優待票_12元")/*  */] OneTime12,
        [Description("販售_優待票_15元")/*  */] OneTime15,
        [Description("販售_普通票_20元")/*  */] OneTime20,
        [Description("販售_普通票_25元")/*  */] OneTime25,
        [Description("販售_普通票_30元")/*  */] OneTime30,
        [Description("販售_自行車票")/*     */] OneTime50,
        [Description("販售_通車紀念套票")/* */] OpenTrafficSetTicket,
        [Description("販售_TPASS")/*        */] TPASS,
        [Description("販售_團體票_張數")/*  */] GroupCount,
        [Description("販售_團體票_人數")/*  */] GroupPeople,
        [Description("販售_團體票_金額")/*  */] GroupAmt,
        [Description("販售_活動1")/*        */] Lanwair1,
        [Description("販售_活動2")/*        */] Lanwair2,
        [Description("AFC取出")/*           */] AfcTakeOut,
        [Description("補票")/*              */] RepairTicket,
        [Description("兌換_一日票")/*       */] ExchangeOneDay,
        [Description("兌換_普通票_20元")/*  */] ExchangeTicket20,
        [Description("兌換_普通票_25元")/*  */] ExchangeTicket25,
        [Description("兌換_普通票_30元")/*  */] ExchangeTicket30,
        [Description("KKDay")/*             */] KKDay,
        [Description("新北幣")/*            */] Ntpc,
        [Description("切票")/*              */] SplitTicker,
        [Description("退費_優待票代墊")/*   */] ConcessionRefund,
        [Description("退費_旅客通知單")/*   */] RefundNotice,
        [Description("退費_售票機")/*       */] VavmRefund,
        [Description("退費_刷卡機")/*       */] PvRefund,
        [Description("退費_信用卡")/*       */] CreditCardRefund,
        [Description("退費_售卡")/*         */] QrCodeRefund,
        [Description("退費_團體票")/*       */] GroupRefund,
    }

    /// <summary>
    /// Client端運行狀態
    /// </summary>
    public enum StatusEnum
    {
        [Description("連接")] Connected,
        [Description("未連接")] Disconnected,
        [Description("NULL")] Validated,
        [Description("NULL")] InSession,
        [Description("完成準備")] Ready,
        [Description("開始")] Start,
        [Description("停止")] Stop,
        [Description("計算")] Calculating,
        [Description("下一個")] Next,
        [Description("回傳")] Returning,
        [Description("輸出比對後資料")] ExportCompareFile,
        [Description("輸出合併後資料")] ExportMegerData,
        [Description("輸出有疑義的資料")] ExportErrorData,
        [Description("輸出一般資料")] ExportTxnData,
        [Description("重新取得卡片交易")] ReGetCardTxn,
        [Description("創建中")] Creating,
        [Description("寫入中")] Inserting,
        [Description("更新中")] Updating,
        [Description("刪除中")] Deleting,
        [Description("讀取中")] Reading,
        [Description("完成")] Finish,
    }

    /// <summary>
    /// 用於Log輸出的詳細程度使用。
    /// </summary>
    public enum ExecutionMode
    {
        [Description("簡易模式")] Simple,
        [Description("普通模式")] Normal,
        [Description("偵錯模式")] Debug,
    }

    /// <summary> 
    /// 語言模式 Num, English, zh-TW 。
    /// </summary>
    public enum Language
    {
        [Description("數字")] Number,
        [Description("英文")] English,
        [Description("繁中")] zhTW,
        [Description("簡中")] zhCN,
    }

    /// <summary>
    /// 控制群組物件用的參數。
    /// </summary>
    public enum CompareGroup
    {
        [Description("第0組")] Zone,
        [Description("第1組")] First,
        [Description("第2組")] Second,
        [Description("第3組")] Third,
        [Description("第4組")] Fourth,
        [Description("第5組")] Fifth,
        [Description("第6組")] Sixth,
        [Description("第7組")] Seventh,
        [Description("第8組")] Eighth,
        [Description("第9組")] Ninth,
        [Description("第10組")] Tenth,
    }

    /// <summary>
    /// 各頁籤
    /// </summary>
    public enum FromTablePage
    {
        [Description("捷運頁籤")] MrtPage,
        [Description("輕軌頁籤")] LrtPage,
        [Description("測試頁籤")] TestPage,
    }

    /// <summary>
    /// 行動支付是否要加運量，要加進站還是出站
    /// </summary>
    public enum MobileEnum
    {
        [Description("兩個都不加")] AddNot,
        [Description("只加出站")] AddExit,
        [Description("只加進站")] AddEntry,
        [Description("兩個都加")] AddBoth,
    }

    /// <summary>
    /// 報表類型 - 1
    /// </summary>
    public enum ReportSwitch_1
    {
        [Description("天期")] Day,
        [Description("車站")] Station,
    }

    /// <summary>
    /// 進站出站
    /// </summary>
    public enum E2
    {
        [Description("進站")] Entry,
        [Description("出站")] Exit,
    }

    /// <summary>
    /// 票種 - 2
    /// </summary>
    public enum Ele_Iss
    {
        [Description("其他")  ] NA,
        [Description("悠遊卡")] ECC,
        [Description("一卡通")] Ipass,
        [Description("愛金卡")] Icash,
        [Description("行動支付")] Mobile,
        [Description("新北捷運")] OwnTicket,
    }

    /// <summary>
    /// 
    /// </summary>
    public enum MobileCard
    {
        [Description("票卡種")] Type,
        [Description("票卡別")] SubType,
    }

    public enum ReportSwitch_2
    {
        [Description("運量")] Volumn,
        [Description("營收")] Amount,
    }

    public enum YearType
    {
        [Description("西元")] AD,
        [Description("民國")] RC,
    }

    public enum FileType
    {
        [Description("csv")] CSV,
        [Description("xlsx")] Excel,
        [Description("docs")] Word,
    }

    public enum MobileVolumn
    {
        [Description("票價扣款")            ] VolumnAmout,
        [Description("出站補扣款")          ] ExitRefine,
        [Description("同站進出時間內不扣款")] OnTime,
        [Description("進站補扣款")          ] EntryRefine,
        [Description("逾時罰款")            ] TimeOut,

    }

    #endregion





}
