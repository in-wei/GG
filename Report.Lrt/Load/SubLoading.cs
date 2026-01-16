using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//--
using System.IO;
//using Report.Lrt.ViewModel;
using UsuallyCommond.MyEnum;
using GlobalCommond;
using GlobalCommond.ViewModel;

namespace Report.Lrt.Load
{
    public class MyCreat
    {
        public static readonly string Path_Sql = $@"./Sql/";

        #region 更新

        public static readonly bool Lrt_UpData_Config_Txt /*-----------------------------------------程式設定檔 */ = false;  //不能更新
        public static readonly bool Lrt_UpData_Config_Line /*------------------------------------------營運路線 */ = true;
        /*-----------------------------------------------------------------------------------------------------------------------*/
        public static readonly bool UpData_Lrt_Report_Csv /*----------------------------------------------------*/ = true;
        public static readonly bool UpData_Config_Sql_List /*-------------------------------------------------- */ = true;
        public static readonly bool UpData_Lrt_Pao_Csv /*------------------------------------------------------ */ = true;
        public static readonly bool UpData_Lrt_StationList_Csv /*--------------------------------------車站代號 */ = true;
        public static readonly bool UpData_MobileCard_Csv /*----------------------------------------------------*/ = true;
        /*-----------------------------------------------------------------------------------------------------------------------*/
        public static readonly bool UpData_Lrt_StationOdList_Csv /*------------------------------------車站計算 */ = true;
        public static readonly bool UpData_Lrt_Subsidy_Csv /*------------------------------------------運價票差 */ = true;
        /*-----------------------------------------------------------------------------------------------------------------------*/
        public static readonly bool UpData_Lrt_Ak_Report_Csv /*------------------------------------------------ */ = true;
        public static readonly bool UpData_Lrt_Ak_Subsidy_Csv /*------------------------------------------------*/ = true;
        public static readonly bool UpData_Lrt_Ak_Sql_Csv /*----------------------------------------------------*/ = true;
        /*-----------------------------------------------------------------------------------------------------------------------*/
        public static readonly bool UpData_Lrt_Dh_Report_Csv /*------------------------------------------------ */ = true;
        public static readonly bool UpData_Lrt_Dh_Subsidy_Csv /*------------------------------------------------*/ = true;
        public static readonly bool UpData_Lrt_Dh_Sql_Csv /*----------------------------------------------------*/ = true;
        /*-----------------------------------------------------------------------------------------------------------------------*/
        public static readonly bool UpData_Sql_Command_Day_Volume /*-------------------------營運資料統計(運量) */ = true;
        public static readonly bool UpData_Sql_Command_Day_Amount /*-------------------------營運資料統計(營收) */ = true;
        public static readonly bool UpData_Sql_Command_Day_EachStation_EachTime /*---------每日全線各站分時運量 */ = true;
        public static readonly bool UpData_Sql_Command_Day_ElectronicTicket /*---------------各家票卡進出站運量 */ = true;
        public static readonly bool UpData_Sql_Command_Day_AllRideList /*--------------------------日運量統計表 */ = true;
        public static readonly bool UpData_Sql_Command_Day_TrafficAmount /*----------------------------每日營收 */ = true;
        public static readonly bool UpData_Sql_Command_Day_OriginDestination /*--------------------每日起迄總表 */ = true;
        public static readonly bool UpData_Sql_Command_Day_EquipAmount /*--------------------------設備營收日報 */ = true;
        public static readonly bool UpData_Sql_Command_Month_AllRideList /*------------------------月運量統計表 */ = true;
        public static readonly bool UpData_Sql_Command_Month_OriginDestination /*------------------每月起訖總表 */ = true;
        public static readonly bool UpData_Sql_Command_Month_ElectronicTicket_Station /*-各家票卡運量月報(車站) */ = true;
        public static readonly bool UpData_Sql_Command_Month_ElectronicTicket_Day /*---- 各家票卡運量月報(天期) */ = true;
        public static readonly bool UpData_Sql_Command_Month_OwnTicketVolume_Station /*--自有票卡運量月報(車站) */ = true;
        public static readonly bool UpData_Sql_Command_Month_OwnTicketVolume_Day /*------自有票卡運量月報(天期) */ = true;
        public static readonly bool UpData_Sql_Command_Month_TrafficAmount_Station /*----------營收月報表(車站) */ = true;
        public static readonly bool UpData_Sql_Command_Month_TrafficAmount_Day /*--------------營收月報表(天期) */ = true;
        public static readonly bool UpData_Sql_Command_Month_EquipAmount_Station /*--------設備營收月報表(車站) */ = true;
        public static readonly bool UpData_Sql_Command_Month_EquipAmount_Day /*------------設備營收月報表(天期) */ = true;


        #endregion

        #region 程式設定

        public static readonly string Lrt_Config_Txt = @"**Mode**
Using Mode,Normal
Export Language,Taiwanese

**Operation**
*UsingLine為使用的各分部 OperationLine為產製營運線*
Using Line,Dh
Operation Line,Test
Operation First Day,2019/02/01
Operation Start Time,05:00:00
Operation End Time,05:00:00

**Server**
Sql Server Ip,192.168.144.8
Sql Catalog,DHLRT_AFC
Sql User Id,sa
Sql User Password,jet..123

**Other**
Excel Applaction Show,0
Thread Use,0
Install MobilePay,1
Auto Pao Excel,1
Test Button,1
Finish Button,18
";
        public static readonly string Lrt_Config_Line = @"運量,代號,路線
Test,Test,測試
Lrt,Dh,淡海
Lrt,Ak,安坑
Mrt,Y,環狀線
";

        public static readonly string Lrt_Station_Csv = @"路線,代碼,代號,名稱,詢問處
Dh,101,V01,紅樹林,1
Dh,102,V02,竿蓁林,0
Dh,103,V03,淡金鄧公,0
Dh,104,V04,淡江大學,0
Dh,105,V05,淡金北新,0
Dh,106,V06,新市一路,0
Dh,107,V07,淡水行政中心,0
Dh,108,V08,濱海義山,0
Dh,109,V09,濱海沙崙,0
Dh,110,V10,淡海新市鎮,0
Dh,111,V11,崁頂,0
Dh,207,V26,淡水漁人碼頭,1
Dh,208,V27,沙崙,0
Dh,209,V28,台北海洋大學,0
Ak,151,K01,雙城,0
Ak,152,K02,玫瑰中國城,0
Ak,153,K03,台北小城,0
Ak,154,K04,耕莘安康院區,0
Ak,155,K05,景文科大,0
Ak,156,K06,安康,0
Ak,157,K07,陽光運動公園,0
Ak,158,K08,新和國小,0
Ak,159,K09,十四張,1
";
        public static readonly string Lrt_Pao_Csv = @"營運,車站,一日票,優待票10元,優待票12元,優待票15元,單程票20元,單程票25元,單程票30元,自行車票,通車紀念套票,行政院通勤月票紀念卡,團體票張數,團體票人數,團體票金額,活動1,活動2,設備取出,補票,兌換單程票20元,兌換單程票25元,兌換單程票30元,一日票(KKDay),一日票(新北幣),切票,兌換一日票,優待票代墊款,退費通知,自動售票機退費,刷卡機退費,信用卡退費,乘車碼退費,退卡,團體票退費
Dh,V01,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1
Dh,V26,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,0,0,1,1,1,1,1,1,1,1,1,1
Ak,K09,1,1,1,0,1,1,0,1,1,1,1,1,1,1,1,1,1,1,1,0,1,1,1,1,1,1,1,1,1,1,1,1
";
        public static readonly string Lrt_Config_Sql_List = @"資料表中文,使用語法1,使用語法2,備註
營運資料統計表(運量),Day_Volume,-,-
營運資料統計表(營收),Day_Amount,-,-
每日全線各站分時運量日報表,Day_EachStation_EachTime,-,-
票卡進出站運量日報表,Day_ElectronicTicket,-,-
日運量統計表,-,-,停用
營收日報表,Day_TrafficAmount,-,-
每日起訖站總表,Day_OriginDestination,-,-
設備營收日報表,Day_EquipAmount,-,-
月運量統計表,-,-,停用
每月起訖總表,Month_OriginDestination,-,與每日起迄站總表共用
各家票卡運量月報表(車站),Month_ElectronicTicket_Station,-,-
各家票卡運量月報表(天期),Month_ElectronicTicket_Day,-,-
自有票卡運量月報表(車站),Month_OwnTicketVolume_Station,-,-
自有票卡運量月報表(天期),Month_OwnTicketVolume_Day,-,-
營收月報表(車站),Month_TrafficAmount_Station,-,-
營收月報表(天期),Month_TrafficAmount_Day,-,-
設備營收月報表(車站),Month_EquipAmount_Station,-,與設備營收日報表共用
設備營收月報表(天期),Month_EquipAmount_Day,-,-
";
        public static readonly string MobileCard_Csv = @"卡片名稱,卡種,卡別
MasterCard,Credit,M
CUP,Credit,C
Discorver,Credit,D
LinePayMoney,QrCode,I
TS QrCode,QrCode,T
";
        public static readonly string Lrt_Subsidy_Csv = @"營運區域,路線代號,起站,訖站,運價,差額
Dh,G,V01,V01,20,0
Dh,G,V01,V02,20,0
Dh,G,V01,V03,20,0
Dh,G,V01,V04,20,10
Dh,G,V01,V05,20,16
Dh,G,V01,V06,20,22
Dh,G,V01,V07,20,31
Dh,G,V01,V08,25,34
Dh,G,V01,V09,25,40
Dh,G,V01,V10,25,48
Dh,G,V01,V11,25,54
Dh,G,V01,V26,25,62
Dh,G,V01,V27,25,58
Dh,G,V01,V28,30,46
Dh,G,V02,V01,20,0
Dh,G,V02,V02,20,0
Dh,G,V02,V03,20,0
Dh,G,V02,V04,20,1
Dh,G,V02,V05,20,7
Dh,G,V02,V06,20,13
Dh,G,V02,V07,20,22
Dh,G,V02,V08,20,30
Dh,G,V02,V09,25,31
Dh,G,V02,V10,25,38
Dh,G,V02,V11,25,45
Dh,G,V02,V26,25,58
Dh,G,V02,V27,25,49
Dh,G,V02,V28,25,37
Dh,G,V03,V01,20,0
Dh,G,V03,V02,20,0
Dh,G,V03,V03,20,0
Dh,G,V03,V04,20,0
Dh,G,V03,V05,20,0
Dh,G,V03,V06,20,3
Dh,G,V03,V07,20,12
Dh,G,V03,V08,20,20
Dh,G,V03,V09,20,26
Dh,G,V03,V10,25,33
Dh,G,V03,V11,25,35
Dh,G,V03,V26,25,47
Dh,G,V03,V27,25,39
Dh,G,V03,V28,25,31
Dh,G,V04,V01,20,10
Dh,G,V04,V02,20,1
Dh,G,V04,V03,20,0
Dh,G,V04,V04,20,0
Dh,G,V04,V05,20,0
Dh,G,V04,V06,20,0
Dh,G,V04,V07,20,1
Dh,G,V04,V08,20,9
Dh,G,V04,V09,20,15
Dh,G,V04,V10,20,23
Dh,G,V04,V11,25,29
Dh,G,V04,V26,25,37
Dh,G,V04,V27,25,33
Dh,G,V04,V28,25,21
Dh,G,V05,V01,20,16
Dh,G,V05,V02,20,7
Dh,G,V05,V03,20,0
Dh,G,V05,V04,20,0
Dh,G,V05,V05,20,0
Dh,G,V05,V06,20,0
Dh,G,V05,V07,20,0
Dh,G,V05,V08,20,3
Dh,G,V05,V09,20,9
Dh,G,V05,V10,20,16
Dh,G,V05,V11,20,23
Dh,G,V05,V26,25,30
Dh,G,V05,V27,25,27
Dh,G,V05,V28,25,14
Dh,G,V06,V01,20,22
Dh,G,V06,V02,20,13
Dh,G,V06,V03,20,3
Dh,G,V06,V04,20,0
Dh,G,V06,V05,20,0
Dh,G,V06,V06,20,0
Dh,G,V06,V07,20,0
Dh,G,V06,V08,20,0
Dh,G,V06,V09,20,4
Dh,G,V06,V10,20,11
Dh,G,V06,V11,20,18
Dh,G,V06,V26,20,30
Dh,G,V06,V27,25,21
Dh,G,V06,V28,25,9
Dh,G,V07,V01,20,31
Dh,G,V07,V02,20,22
Dh,G,V07,V03,20,12
Dh,G,V07,V04,20,1
Dh,G,V07,V05,20,0
Dh,G,V07,V06,20,0
Dh,G,V07,V07,20,0
Dh,G,V07,V08,20,0
Dh,G,V07,V09,20,0
Dh,G,V07,V10,20,1
Dh,G,V07,V11,20,8
Dh,G,V07,V26,20,21
Dh,G,V07,V27,20,12
Dh,G,V07,V28,25,0
Dh,G,V08,V01,25,34
Dh,G,V08,V02,20,30
Dh,G,V08,V03,20,20
Dh,G,V08,V04,20,9
Dh,G,V08,V05,20,3
Dh,G,V08,V06,20,0
Dh,G,V08,V07,20,0
Dh,G,V08,V08,20,0
Dh,G,V08,V09,20,0
Dh,G,V08,V10,20,0
Dh,G,V08,V11,20,0
Dh,G,V08,V26,20,13
Dh,G,V08,V27,20,4
Dh,G,V08,V28,20,0
Dh,G,V09,V01,25,40
Dh,G,V09,V02,25,31
Dh,G,V09,V03,20,26
Dh,G,V09,V04,20,15
Dh,G,V09,V05,20,9
Dh,G,V09,V06,20,4
Dh,G,V09,V07,20,0
Dh,G,V09,V08,20,0
Dh,G,V09,V09,20,0
Dh,G,V09,V10,20,0
Dh,G,V09,V11,20,0
Dh,G,V09,V26,20,6
Dh,G,V09,V27,20,0
Dh,G,V09,V28,20,0
Dh,G,V10,V01,25,47
Dh,G,V10,V02,25,38
Dh,G,V10,V03,25,33
Dh,G,V10,V04,20,22
Dh,G,V10,V05,20,16
Dh,G,V10,V06,20,11
Dh,G,V10,V07,20,1
Dh,G,V10,V08,20,0
Dh,G,V10,V09,20,0
Dh,G,V10,V10,20,0
Dh,G,V10,V11,20,0
Dh,G,V10,V26,20,13
Dh,G,V10,V27,20,5
Dh,G,V10,V28,20,0
Dh,G,V11,V01,25,54
Dh,G,V11,V02,25,45
Dh,G,V11,V03,25,35
Dh,G,V11,V04,25,29
Dh,G,V11,V05,20,23
Dh,G,V11,V06,20,17
Dh,G,V11,V07,20,8
Dh,G,V11,V08,20,0
Dh,G,V11,V09,20,0
Dh,G,V11,V10,20,0
Dh,G,V11,V11,20,0
Dh,G,V11,V26,20,20
Dh,G,V11,V27,20,11
Dh,G,V11,V28,20,0
Dh,BL,V26,V01,25,62
Dh,BL,V26,V02,25,58
Dh,BL,V26,V03,25,48
Dh,BL,V26,V04,25,37
Dh,BL,V26,V05,25,31
Dh,BL,V26,V06,20,30
Dh,BL,V26,V07,20,21
Dh,BL,V26,V08,20,13
Dh,BL,V26,V09,20,6
Dh,BL,V26,V10,20,14
Dh,BL,V26,V11,20,20
Dh,BL,V26,V26,20,0
Dh,BL,V26,V27,20,0
Dh,BL,V26,V28,20,1
Dh,BL,V27,V01,25,58
Dh,BL,V27,V02,25,49
Dh,BL,V27,V03,25,39
Dh,BL,V27,V04,25,33
Dh,BL,V27,V05,25,27
Dh,BL,V27,V06,25,21
Dh,BL,V27,V07,20,12
Dh,BL,V27,V08,20,4
Dh,BL,V27,V09,20,0
Dh,BL,V27,V10,20,5
Dh,BL,V27,V11,20,12
Dh,BL,V27,V26,20,0
Dh,BL,V27,V27,20,0
Dh,BL,V27,V28,20,0
Dh,BL,V28,V01,30,46
Dh,BL,V28,V02,25,37
Dh,BL,V28,V03,25,32
Dh,BL,V28,V04,25,21
Dh,BL,V28,V05,25,14
Dh,BL,V28,V06,25,9
Dh,BL,V28,V07,25,0
Dh,BL,V28,V08,20,0
Dh,BL,V28,V09,20,0
Dh,BL,V28,V10,20,0
Dh,BL,V28,V11,20,0
Dh,BL,V28,V26,20,1
Dh,BL,V28,V27,20,0
Dh,BL,V28,V28,20,0
Ak,BN,K01,K01,20,0
Ak,BN,K01,K02,20,0
Ak,BN,K01,K03,20,0
Ak,BN,K01,K04,20,0
Ak,BN,K01,K05,20,7
Ak,BN,K01,K06,20,22
Ak,BN,K01,K07,20,31
Ak,BN,K01,K08,25,41
Ak,BN,K01,K09,25,53
Ak,BN,K02,K01,20,0
Ak,BN,K02,K02,20,0
Ak,BN,K02,K03,20,0
Ak,BN,K02,K04,20,0
Ak,BN,K02,K05,20,0
Ak,BN,K02,K06,20,15
Ak,BN,K02,K07,20,24
Ak,BN,K02,K08,20,35
Ak,BN,K02,K09,25,46
Ak,BN,K03,K01,20,0
Ak,BN,K03,K02,20,0
Ak,BN,K03,K03,20,0
Ak,BN,K03,K04,20,0
Ak,BN,K03,K05,20,0
Ak,BN,K03,K06,20,8
Ak,BN,K03,K07,20,17
Ak,BN,K03,K08,20,32
Ak,BN,K03,K09,20,39
Ak,BN,K04,K01,20,1
Ak,BN,K04,K02,20,0
Ak,BN,K04,K03,20,0
Ak,BN,K04,K04,20,0
Ak,BN,K04,K05,20,0
Ak,BN,K04,K06,20,2
Ak,BN,K04,K07,20,11
Ak,BN,K04,K08,20,27
Ak,BN,K04,K09,20,33
Ak,BN,K05,K01,20,7
Ak,BN,K05,K02,20,0
Ak,BN,K05,K03,20,0
Ak,BN,K05,K04,20,0
Ak,BN,K05,K05,20,0
Ak,BN,K05,K06,20,0
Ak,BN,K05,K07,20,4
Ak,BN,K05,K08,20,20
Ak,BN,K05,K09,20,31
Ak,BN,K06,K01,20,22
Ak,BN,K06,K02,20,15
Ak,BN,K06,K03,20,8
Ak,BN,K06,K04,20,1
Ak,BN,K06,K05,20,0
Ak,BN,K06,K06,20,0
Ak,BN,K06,K07,20,0
Ak,BN,K06,K08,20,4
Ak,BN,K06,K09,20,16
Ak,BN,K07,K01,20,31
Ak,BN,K07,K02,20,24
Ak,BN,K07,K03,20,17
Ak,BN,K07,K04,20,10
Ak,BN,K07,K05,20,4
Ak,BN,K07,K06,20,0
Ak,BN,K07,K07,20,0
Ak,BN,K07,K08,20,0
Ak,BN,K07,K09,20,7
Ak,BN,K08,K01,25,41
Ak,BN,K08,K02,20,35
Ak,BN,K08,K03,20,32
Ak,BN,K08,K04,20,26
Ak,BN,K08,K05,20,20
Ak,BN,K08,K06,20,4
Ak,BN,K08,K07,20,0
Ak,BN,K08,K08,20,0
Ak,BN,K08,K09,20,0
Ak,BN,K09,K01,25,53
Ak,BN,K09,K02,25,46
Ak,BN,K09,K03,20,39
Ak,BN,K09,K04,20,32
Ak,BN,K09,K05,20,31
Ak,BN,K09,K06,20,16
Ak,BN,K09,K07,20,7
Ak,BN,K09,K08,20,0
Ak,BN,K09,K09,20,0
";
        public static readonly string Lrt_Report_Csv = @"路線,報表中文名,報表英文名,產製日期欄(英),產製日期行(數),起始欄(英),起始行(數),結束欄(英),結束行(數),斷點行(數),斷點行數,斷點欄(英),斷點欄數,電子支付起始欄(英),自有票起始欄(英)
Dh,營運資料統計表(運量),Day_Volume,-,-,B,8,U,8,-,-,-,-,J,L
Dh,營運資料統計表(營收),Day_Amount,-,-,B,8,Y,8,-,-,-,-,H,J
Dh,每日全線各站分時運量日報表,Day_EachStation_EachTime,C,4,D,8,AE,32,-,-,-,-,D,D
Dh,票卡進出站運量日報表,Day_ElectronicTicket,D,4,G,9,AD,22,-,-,-,-,Q,U
Dh,日運量統計表,-,-,-,-,-,-,-,-,-,-,-,-,-
Dh,營收日報表,Day_TrafficAmount,D,4,F,9,V,22,-,-,-,-,I,K
Dh,每日起訖站總表,Day_OriginDestination,D,4,D,7,Q,22,-,-,-,-,-,-
Dh,設備營收日報表,Day_EquipAmount,D,4,F,10,AU,23,-,-,AC,3,AG,-
Dh,月運量統計表,-,-,-,-,-,-,-,-,-,-,-,-,-
Dh,每月起訖總表,Month_OriginDestination,D,4,D,7,Q,20,-,-,-,-,-,-
Dh,各家票卡運量月報表(車站),Month_ElectronicTicket_Station,D,4,F,9,Z,22,-,-,-,-,V,-
Dh,各家票卡運量月報表(天期),Month_ElectronicTicket_Day,D,4,F,9,Z,39,-,-,-,-,V,-
Dh,自有票卡運量月報表(車站),Month_OwnTicketVolume_Station,D,4,F,9,S,22,-,-,-,-,-,K
Dh,自有票卡運量月報表(天期),Month_OwnTicketVolume_Day,D,4,F,9,Z,38,-,-,-,-,-,K
Dh,營收月報表(車站),Month_TrafficAmount_Station,D,4,F,9,V,22,-,-,-,-,I,K
Dh,營收月報表(天期),Month_TrafficAmount_Day,D,4,F,9,V,38,-,-,-,-,I,K
Dh,設備營收月報表(車站),Month_EquipAmount_Station,D,4,F,10,AU,23,-,-,AC,3,AG,-
Dh,設備營收月報表(天期),Month_EquipAmount_Day,D,4,F,10,AU,39,-,-,AC,3,AG,-
Ak,營運資料統計表(運量),Day_Volume,-,-,B,8,U,8,-,-,-,-,J,L
Ak,營運資料統計表(營收),Day_Amount,-,-,B,8,Y,8,-,-,-,-,H,J
Ak,每日全線各站分時運量日報表,Day_EachStation_EachTime,C,4,AJ,8,BA,32,-,-,-,-,AJ,AJ
Ak,票卡進出站運量日報表,Day_ElectronicTicket,D,4,G,24,AD,32,-,-,-,-,Q,U
Ak,日運量統計表,-,-,-,-,-,-,-,-,-,-,-,-,-
Ak,營收日報表,Day_TrafficAmount,D,4,F,24,V,32,-,-,-,-,I,K
Ak,每日起訖站總表,Day_OriginDestination,D,4,D,23,L,31,-,-,-,-,-,-
Ak,設備營收日報表,Day_EquipAmount,D,4,F,25,AU,33,-,-,AC,3,AG,-
Ak,月運量統計表,-,-,-,-,-,-,-,-,-,-,-,-,-
Ak,每月起訖總表,Month_OriginDestination,D,4,D,23,L,31,-,-,-,-,-,-
Ak,各家票卡運量月報表(車站),Month_ElectronicTicket_Station,D,4,F,24,Z,32,-,-,-,-,V,-
Ak,各家票卡運量月報表(天期),Month_ElectronicTicket_Day,D,4,F,9,Z,38,-,-,-,-,V,-
Ak,自有票卡運量月報表(車站),Month_OwnTicketVolume_Station,D,4,F,24,Z,32,-,-,-,-,-,K
Ak,自有票卡運量月報表(天期),Month_OwnTicketVolume_Day,D,4,F,9,Z,38,-,-,-,-,-,K
Ak,營收月報表(車站),Month_TrafficAmount_Station,D,4,F,24,V,32,-,-,-,-,I,K
Ak,營收月報表(天期),Month_TrafficAmount_Day,D,4,F,9,V,38,-,-,-,-,I,K
Ak,設備營收月報表(車站),Month_EquipAmount_Station,D,4,F,25,AU,33,-,-,AC,3,AG,-
Ak,設備營收月報表(天期),Month_EquipAmount_Day,D,4,F,10,AU,39,-,-,AC,3,AG,-
";

        #endregion

        #region SQL

        public static readonly string Sql_Command_Day_Volume = @"--營運資料統計表(運量)
DECLARE @StartStation VARCHAR(10)    --該路線車站起點
DECLARE @EndStation VARCHAR(10)        --該路線車站終點
DECLARE @Openstardate VARCHAR(20)    --搜尋時間開始
DECLARE @Openenddate VARCHAR(20)    --搜尋時間結束
DECLARE @Monthstardate VARCHAR(20)    --開始的月份
DECLARE @Monthenddate VARCHAR(20)    --結束的月份
DECLARE @DateSet VARCHAR(20)        --計算月份用
DECLARE @RunDate VARCHAR(20)        --執行的月份

SET @StartStation = '{0}' 
SET @EndStation = '{1}'
SET @Openstardate = '{2}'
SET @Openenddate = '{3}'
SET @DateSet = CONCAT(DATEPART(YEAR,@Openstardate),'-',DATEPART(MONTH,@Openstardate) , '-01 0' , DATEPART(HOUR,@Openstardate) , ':0' , DATEPART(MINUTE,@Openstardate) , ':0' , DATEPART(SECOND,@Openstardate))
SET @Monthenddate = DATEADD(MONTH,1,@DateSet)
SET @Monthstardate = DATEADD(MONTH,-1,@Monthenddate)
SET @RunDate = @Monthstardate

SELECT LocationId AS mLoc
INTO #mTempLocation
FROM Parm081_LocList
WHERE LocationId >= @StartStation AND LocationId <= @EndStation AND CurrentType IN (1)

SELECT *
INTO #mTempExit
FROM Txn_Exit JOIN #mTempLocation ON LocationId = mLoc
WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate

SELECT *
INTO #mTempTicket
FROM Txn_SellSpecialTicket JOIN #mTempLocation ON LocationId = mLoc
WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate

SELECT *
INTO #mTempRefund
FROM Txn_RefundSpeTkt JOIN #mTempLocation ON LocationId = mLoc
WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate

SELECT 
	(ISNULL(All_Exit,0) - ISNULL(Welfare_Exit,0) - ISNULL(Student_Exit,0) - ISNULL(OneDayEle,0) - ISNULL(Officeman,0) - ISNULL(All_Pass_Exit,0)) AS Electron,
	ISNULL(Student_Exit,0) AS Student_Exit,
	ISNULL(Welfare_Exit,0) AS Welfare_Exit,
	ISNULL(ISNULL(All_Pass_Exit,0) - ISNULL(All_Pass_Student_Exit,0),0) AS All_Pass_Exit,
	ISNULL(All_Pass_Student_Exit,0) AS All_Pass_Student_Exit,
	ISNULL(CreditCard,0) AS CreditCard,
	ISNULL(QrCode,0) AS QrCode,
	ISNULL(ISNULL(SOneTicket_Normal,0) - ISNULL(SOneTicket_Normal_Refund,0),0) AS SOneTicket_Normal,
	ISNULL(ISNULL(SOneTicket_Walf,0) - ISNULL(SOneTicket_Walf_Refund,0),0) AS SOneTicket_Walf,
	ISNULL(ISNULL(SOneTicket_Bike,0) - ISNULL(SOneTicket_Bike_Refund,0),0) AS SOneTicket_Bike,
	ISNULL(OneDayTicket,0) AS OneDayTicket,
	ISNULL(OneDayEle,0) AS OneDayEle,
	ISNULL(OneDayTicket_1,0) AS OneDayTicket_1,
	ISNULL(OneDayTicket_2,0) AS OneDayTicket_2,
	ISNULL(GroupTicketCount,0) AS GroupTicketCount,
	ISNULL(GroupTicketPeople,0) AS GroupTicketPeople,
	ISNULL(Officeman,0) AS Officeman

FROM
          (SELECT COUNT(*) AS All_Exit                 FROM #mTempExit                                                                            ) AS All_Exit
LEFT JOIN (SELECT COUNT(*) AS Welfare_Exit             FROM #mTempExit   WHERE TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 
        AND (
		       (IssueCode IN (2) AND PersonalProfile IN (1,2,3,4)) 
			OR (IssueCode IN (2) AND PersonalProfile IN (8) AND AreaCode IN (1,2)) 
			OR (IssueCode IN (9) AND IdentityType IN (3,4,6,5)) 
			OR (IssueCode IN (9) AND IdentityType IN (2) AND AreaCode IN (24,25,30,34,38,42,46,50,54,58,62,66,70,74,78,82,86,90,94,98,102,106,110,114))
		)
) AS Welfare_Exit ON 1=1
LEFT JOIN (SELECT COUNT(*) AS Student_Exit            FROM #mTempExit   WHERE TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0
        AND (
               (IssueCode IN (2) AND PersonalProfile in (5))
            OR (IssueCode IN (9) AND IdentityType IN (2) AND CardType IN (7) AND AreaCode NOT IN (24,25,30,34,38,42,46,50,54,58,62,66,70,74,78,82,86,90,94,98,102,106,110,114))
            OR (IssueCode IN (11) AND CardType IN (2))
		)
) AS Student_Exit ON 1=1
LEFT JOIN (SELECT COUNT(*) AS All_Pass_Exit            FROM #mTempExit   WHERE TxnType IN (28,30) AND IssueCode IN (2)                           ) AS All_Pass_Exit            ON 1=1
LEFT JOIN (SELECT COUNT(*) AS All_Pass_Student_Exit    FROM #mTempExit   WHERE TxnType IN (28,30) AND IssueCode IN (2) AND PersonalProfile in (5)) AS All_Pass_Student_Exit    ON 1=1
LEFT JOIN (SELECT 0 AS CreditCard                                                                                                                ) AS CreditCard               ON 1=1
LEFT JOIN (SELECT 0 AS QrCode                                                                                                                    ) AS QrCode                   ON 1=1
LEFT JOIN (SELECT COUNT(*) AS SOneTicket_Normal        FROM #mTempTicket WHERE CardType IN (1)                                                   ) AS SOneTicket_Normal        ON 1=1
LEFT JOIN (SELECT COUNT(*) AS SOneTicket_Normal_Refund FROM #mTempRefund WHERE CardType IN (1)                                                   ) AS SOneTicket_Normal_Refund ON 1=1
LEFT JOIN (SELECT COUNT(*) AS SOneTicket_Walf          FROM #mTempTicket WHERE CardType IN (3)                                                   ) AS SOneTicket_Walf          ON 1=1
LEFT JOIN (SELECT COUNT(*) AS SOneTicket_Walf_Refund   FROM #mTempRefund WHERE CardType IN (3)                                                   ) AS SOneTicket_Walf_Refund   ON 1=1
LEFT JOIN (SELECT COUNT(*) AS SOneTicket_Bike          FROM #mTempTicket WHERE CardType IN (2)                                                   ) AS SOneTicket_Bike          ON 1=1
LEFT JOIN (SELECT COUNT(*) AS SOneTicket_Bike_Refund   FROM #mTempRefund WHERE CardType IN (2)                                                   ) AS SOneTicket_Bike_Refund   ON 1=1
LEFT JOIN (SELECT 0 AS OneDayTicket                                                                                                              ) AS OneDayTicket             ON 1=1
LEFT JOIN (SELECT COUNT(*) AS OneDayEle                FROM #mTempExit   WHERE TxnType IN (24) AND PeriodCode IN (1) /*AND IssueCode IN (9)*/    ) AS OneDayEle                ON 1=1
LEFT JOIN (SELECT 0 AS OneDayTicket_1                                                                                                            ) AS OneDayTicket_1           ON 1=1
LEFT JOIN (SELECT 0 AS OneDayTicket_2                                                                                                            ) AS OneDayTicket_2           ON 1=1
LEFT JOIN (SELECT 0 AS GroupTicketCount, 0 AS GroupTicketPeople                                                                                  ) AS GroupTicket              ON 1=1
LEFT JOIN (SELECT COUNT(*) AS Officeman                FROM #mTempExit   WHERE TxnType IN (24) AND PeriodCode IN (17) AND IssueCode IN (2,9,11)  ) AS Officeman                ON 1=1

DROP TABLE #mTempRefund
DROP TABLE #mTempTicket
DROP TABLE #mTempExit
DROP TABLE #mTempLocation";
        public static readonly string Sql_Command_Day_Amount = @"--營運資料統計表(營收)
DECLARE @StartStation VARCHAR(10)	--該路線車站起點
DECLARE @EndStation VARCHAR(10)		--該路線車站終點
DECLARE @Openstardate VARCHAR(20)	--搜尋時間開始
DECLARE @Openenddate VARCHAR(20)	--搜尋時間結束
DECLARE @Monthstardate VARCHAR(20)	--開始的月份
DECLARE @Monthenddate VARCHAR(20)	--結束的月份
DECLARE @DateSet VARCHAR(20)		--計算月份用
DECLARE @RunDate VARCHAR(20)		--執行的月份

SET @StartStation = '{0}' 
SET @EndStation = '{1}'
SET @Openstardate = '{2}'
SET @Openenddate = '{3}'
SET @DateSet = CONCAT(DATEPART(YEAR,@Openstardate),'-',DATEPART(MONTH,@Openstardate) , '-01 0' , DATEPART(HOUR,@Openstardate) , ':0' , DATEPART(MINUTE,@Openstardate) , ':0' , DATEPART(SECOND,@Openstardate))
SET @Monthenddate = DATEADD(MONTH,1,@DateSet)
SET @Monthstardate = DATEADD(MONTH,-1,@Monthenddate)
SET @RunDate = @Monthstardate

SELECT LocationId as mLoc
INTO #mTempLocation
FROM Parm081_LocList
WHERE LocationId >= @StartStation AND LocationId <= @EndStation AND CurrentType IN (1)

SELECT (ElectronAmt - StudentAmt - ConcessionTxn + TxnAmtCsc + TxnAmtDed) as NormalAmt, StudentAmt, ConcessionTxn, '0' as CerditAmt, '0' as QrCode, (TxnAmt1 - TxnAmt1Back) as SOneTimeN, (TxnAmt3 - TxnAmt3Back) as SOneTimeC, (TxnAmt2 - TxnAmt2Back) as SOneTimeB, '0' as OneDay_1, '0' as OneDay_2, '0' as OneDay_3, '0' as OneDay_4, '0' as GroupTicket, '0' as OpenTransfare, '0' as RepairTicker, ConcessionWelfare, ConcessionWp, '0' as SoneTicket, TrnAmt, AllPassEle_Amt
FROM
    (
        SELECT ISNULL(SUM(TxnAmt),0) AS ElectronAmt
        FROM Txn_Exit JOIN #mTempLocation on LocationId = mLoc WHERE TxnDT > @Openstardate AND TxnDT < @Openenddate 
    ) AS AllElectron_Exit 
    LEFT JOIN (
        SELECT ISNULL(SUM(TrnAmt),0) AS TrnAmt
        FROM Txn_Exit JOIN #mTempLocation on LocationId = mLoc WHERE TxnDT > @Openstardate AND TxnDT < @Openenddate AND LastTrnFlag in (35) AND TrnAmt IN (5)
    ) AS YouBickTrn ON 1=1 
    LEFT JOIN (
        SELECT ISNULL(SUM(TxnAmt),0) AS StudentAmt
        FROM Txn_Exit JOIN #mTempLocation on LocationId = mLoc 
        WHERE TxnDT > @Openstardate AND TxnDT < @Openenddate 
            AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 
            AND
            (
                   (IssueCode IN (2) AND PersonalProfile IN (5)) 
                OR (IssueCode IN (9) AND IdentityType IN (2) AND CardType IN (7) AND AreaCode not IN (82,78)) 
                OR (IssueCode IN (11) AND CardType IN (2))
            )
    ) AS Student_Exit ON 1=1 
    LEFT JOIN (
        SELECT 
              ISNULL(SUM(TxnAmt),0) AS ConcessionTxn
            , ISNULL(SUM(WelfareAmt),0) AS ConcessionWelfare
            , ISNULL(SUM(TxnWP),0) AS ConcessionWp
        FROM Txn_Exit JOIN #mTempLocation on LocationId = mLoc 
        WHERE TxnDT > @Openstardate AND TxnDT < @Openenddate 
            AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0  
            AND
            (
                   (IssueCode IN (2) AND PersonalProfile IN (1,2,3,4)) 
                OR (IssueCode IN (2) AND PersonalProfile IN (8) AND AreaCode IN (1,2)) 
                OR (IssueCode IN (9) AND IdentityType IN (2) AND AreaCode IN (24,25,30,34,38,42,46,50,54,58,62,66,70,74,78,82,86,90,94,98,102,106,110,114))
                OR (IssueCode IN (9) AND IdentityType IN (2) AND AreaCode IN (119,118)) 
                OR (IssueCode IN (9) AND IdentityType IN (3,4,6,5)) 
            )
    ) AS Concession_Exit ON 1=1 
    LEFT JOIN (
        SELECT AllPassEle_num, AllPassEle_Amt, AllPassEle_Dis, AllPassStudent_num, AllPassStudent_Amt, AllPassStudent_Dis
        FROM 
        (
            SELECT 
                txnyr, txnmh, txndy, COUNT(*) AS AllPassEle_num, SUM(all_pass_Amt) AS AllPassEle_Amt, SUM(all_discount_Amt) AS AllPassEle_Dis
            FROM (
                SELECT 
                      DATEPART(YEAR, T2.OperationDate) AS txnyr
                    , DATEPART(MONTH, T2.OperationDate) AS txnmh
                    , DATEPART(Day, T2.OperationDate) AS txndy 
                    , CONVERT(INT, CONVERT(VARBINARY, SUBSTRING(BodyData, 125, 4), 2)) AS all_pass_Amt
                    , CONVERT(INT, CONVERT(VARBINARY, SUBSTRING(BodyData, 129, 4), 2)) AS all_discount_Amt
                    , TrnAmt
                    , T2.LocationId
                FROM 
                    ECCTxn_Interface AS T1 JOIN Txn_Exit AS T2 ON T1.CardSN = T2.CardSN AND T1.TxnDT = T2.TxnDT JOIN #mTempLocation on T1.LocationId = mLoc
                WHERE T1.TxnDT > @Openstardate AND T1.TxnDT < @Openenddate 
                    AND T1.TxnType IN (4) 
                    AND NOT 
                    (
                        (
                                T1.IssueCode IN (2) 
                            AND PersonalProfile IN (5)
                        ) OR (
                                T1.IssueCode IN (9) 
                            AND IdentityType IN (2) 
                            AND CardType IN (7)
                        ) OR (
                                T1.IssueCode IN (11) 
                            AND CardType IN (2)
                        )
                    )
            ) AS TT GROUP BY txnyr,txnmh,txndy 
        ) AS AllPassEle
    LEFT JOIN (
            SELECT 
                txnyr, txnmh, txndy, Count(*) AS AllPassStudent_num, SUM(all_pass_Amt) AS AllPassStudent_Amt , SUM(all_discount_Amt) AS AllPassStudent_Dis
            FROM 
            (
                SELECT 
                      T2.LocationId, DATEPART(YEAR, T2.OperationDate) AS txnyr
                    , DATEPART(MONTH, T2.OperationDate) AS txnmh
                    , DATEPART(Day, T2.OperationDate) AS txndy 
                    , CONVERT(INT, CONVERT(VARBINARY, SUBSTRING(BodyData, 125, 4), 2)) AS all_pass_Amt
                    , CONVERT(INT, CONVERT(VARBINARY, SUBSTRING(BodyData, 129, 4), 2)) AS all_discount_Amt
                    , TrnAmt
                FROM 
                    ECCTxn_Interface AS T1 JOIN Txn_Exit AS T2 ON T1.CardSN = T2.CardSN AND T1.TxnDT = T2.TxnDT JOIN #mTempLocation on T1.LocationId = mLoc
                WHERE T1.TxnDT > @Openstardate AND T1.TxnDT < @Openenddate 
                    AND T1.TxnType IN (4) 
                    AND 
                    (
                           (T1.IssueCode IN (2) AND PersonalProfile IN (5)) 
                        OR (T1.IssueCode IN (9) AND IdentityType IN (2) AND CardType IN (7)) 
                        OR (T1.IssueCode IN (11) AND CardType IN (2))
                    )
            ) AS TT GROUP BY txnyr,txnmh,txndy 
        ) AS AllPassStudent ON AllPassEle.txnyr = AllPassStudent.txnyr AND AllPassEle.txnmh = AllPassStudent.txnmh AND AllPassEle.txndy = AllPassStudent.txndy
    ) AS AllPAss ON 1=1 
    LEFT JOIN (
        SELECT ISNULL(SUM(TxnAmt),0) AS TxnAmtCsc
        FROM Txn_ExcessByCSC JOIN #mTempLocation on LocationId = mLoc WHERE TxnDT > @Openstardate AND TxnDT < @Openenddate 
    ) AS IcashCsc ON 1=1 
    LEFT JOIN (
        SELECT ISNULL(SUM(TxnAmt),0) AS TxnAmtDed
        FROM Txn_ExceptDeduct JOIN #mTempLocation on LocationId = mLoc WHERE TxnDT > @Openstardate AND TxnDT < @Openenddate 
    ) AS IcashDed ON 1=1 
    LEFT JOIN (
        SELECT ISNULL(SUM(TxnAmt),0) AS TxnAmt1
        FROM Txn_SellSpecialTicket JOIN #mTempLocation on LocationId = mLoc WHERE TxnDT > @Openstardate AND TxnDT < @Openenddate 
        AND CardType IN (1) AND EquipType NOT IN (6)
    ) AS sOneTime_1 ON 1=1 
    LEFT JOIN (
        SELECT ISNULL(SUM(TxnAmt),0) AS TxnAmt3
        FROM Txn_SellSpecialTicket JOIN #mTempLocation on LocationId = mLoc WHERE TxnDT > @Openstardate AND TxnDT < @Openenddate 
        AND CardType IN (3) AND EquipType NOT IN (6)
    ) AS sOneTime_3 ON 1=1 
    LEFT JOIN (
        SELECT ISNULL(SUM(TxnAmt),0) AS TxnAmt2
        FROM Txn_SellSpecialTicket JOIN #mTempLocation on LocationId = mLoc WHERE TxnDT > @Openstardate AND TxnDT < @Openenddate 
        AND CardType IN (2) AND EquipType NOT IN (6)
    ) AS sOneTime_2 ON 1=1 
    LEFT JOIN (
        SELECT ISNULL(SUM(TxnAmt),0) AS TxnAmt1Back
        FROM Txn_RefundSpeTkt JOIN #mTempLocation on LocationId = mLoc WHERE TxnDT > @Openstardate AND TxnDT < @Openenddate 
        AND CardType IN (1)
    ) AS sOneTime_1Back ON 1=1 
    LEFT JOIN (
        SELECT ISNULL(SUM(TxnAmt),0) AS TxnAmt3Back
        FROM Txn_RefundSpeTkt JOIN #mTempLocation on LocationId = mLoc WHERE TxnDT > @Openstardate AND TxnDT < @Openenddate 
        AND CardType IN (3)
    ) AS sOneTime_3Back ON 1=1 
    LEFT JOIN (
        SELECT ISNULL(SUM(TxnAmt),0) AS TxnAmt2Back
        FROM Txn_RefundSpeTkt JOIN #mTempLocation on LocationId = mLoc WHERE TxnDT > @Openstardate AND TxnDT < @Openenddate 
        AND CardType IN (2)
    ) AS sOneTime_2Back ON 1=1 
    LEFT JOIN (
        SELECT ISNULL(SUM(TxnAmt),0) AS TxnAmtPta
        FROM Txn_SellSpecialTicket JOIN #mTempLocation on LocationId = mLoc WHERE TxnDT > @Openstardate AND TxnDT < @Openenddate 
        AND EquipType IN (6)
    ) AS sOneTime_Pta ON 1=1 

DROP TABLE #mTempLocation
";
        public static readonly string Sql_Command_Day_EachStation_EachTime = @"--每日全線各站分時運量日報表
DECLARE @StartStation VARCHAR(10)	--該路線車站起點
DECLARE @EndStation VARCHAR(10)		--該路線車站終點
DECLARE @Openstardate VARCHAR(20)	--搜尋時間開始
DECLARE @Openenddate VARCHAR(20)	--搜尋時間結束
DECLARE @Monthstardate VARCHAR(20)	--開始的月份
DECLARE @Monthenddate VARCHAR(20)	--結束的月份
DECLARE @DateSet VARCHAR(20)		--計算月份用
DECLARE @RunDate VARCHAR(20)		--執行的月份

SET @StartStation = '{0}' 
SET @EndStation = '{1}'
SET @Openstardate = '{2}'
SET @Openenddate = '{3}'
SET @DateSet = CONCAT(DATEPART(YEAR,@Openstardate),'-',DATEPART(MONTH,@Openstardate) , '-01 0' , DATEPART(HOUR,@Openstardate) , ':0' , DATEPART(MINUTE,@Openstardate) , ':0' , DATEPART(SECOND,@Openstardate))
SET @Monthenddate = DATEADD(MONTH,1,@DateSet)
SET @Monthstardate = DATEADD(MONTH,-1,@Monthenddate)
SET @RunDate = @Monthstardate

SELECT LocationId as mLoc
INTO #mTempLocation
FROM Parm081_LocList
WHERE LocationId >= @StartStation AND LocationId <= @EndStation AND CurrentType IN (1)

      SELECT txnyr, txnmh, txndy, txnhr, 'Entry'            as mCheck, '0' AS mCheckOffset, '+' as mCheckIO, LocationId     as LocationId, COUNT(LocationId)    as num FROM (SELECT DATEPART(YEAR, TxnDT) as txnyr, DATEPART(MONTH, TxnDT) as txnmh, DATEPART(DAY, TxnDT) as txndy, DATEPART(Hour, TxnDT) AS txnhr, LocationId      FROM Txn_Entry              JOIN #mTempLocation ON LocationId   = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT<@Openenddate) AS T2  GROUP BY txnyr, txnmh, txndy, txnhr, LocationId 
UNION SELECT txnyr, txnmh, txndy, txnhr, 'Exit'             as mCheck, '1' AS mCheckOffset, '+' as mCheckIO, LocationId     as LocationId, COUNT(LocationId)    as num FROM (SELECT DATEPART(YEAR, TxnDT) as txnyr, DATEPART(MONTH, TxnDT) as txnmh, DATEPART(DAY, TxnDT) as txndy, DATEPART(Hour, TxnDT) as txnhr, LocationId      FROM Txn_Exit               JOIN #mTempLocation ON LocationId   = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT<@Openenddate) AS T2  GROUP BY txnyr, txnmh, txndy, txnhr, LocationId 
UNION SELECT txnyr, txnmh, txndy, txnhr, 'SOneEntry'        as mCheck, '0' AS mCheckOffset, '+' as mCheckIO, StartLoc       as LocationId, COUNT(StartLoc)      as num FROM (SELECT DATEPART(YEAR, TxnDT) as txnyr, DATEPART(MONTH, TxnDT) as txnmh, DATEPART(DAY, TxnDT) as txndy, DATEPART(Hour, TxnDT) as txnhr, StartLoc        FROM Txn_SellSpecialTicket  JOIN #mTempLocation ON StartLoc     = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT<@Openenddate) AS T2  GROUP BY txnyr, txnmh, txndy, txnhr, StartLoc 
UNION SELECT txnyr, txnmh, txndy, txnhr, 'SOneExit'         as mCheck, '1' AS mCheckOffset, '+' as mCheckIO, EndLoc         as LocationId, COUNT(EndLoc)        as num FROM (SELECT DATEPART(YEAR, TxnDT) as txnyr, DATEPART(MONTH, TxnDT) as txnmh, DATEPART(DAY, TxnDT) as txndy, DATEPART(Hour, TxnDT) as txnhr, EndLoc          FROM Txn_SellSpecialTicket  JOIN #mTempLocation ON StartLoc     = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT<@Openenddate) AS T2  GROUP BY txnyr, txnmh, txndy, txnhr, EndLoc 
UNION SELECT txnyr, txnmh, txndy, txnhr, 'SOneEntryBack'    as mCheck, '0' AS mCheckOffset, '-' as mCheckIO, SaleStartLoc   as LocationId, COUNT(SaleStartLoc)  as num FROM (SELECT DATEPART(YEAR, TxnDT) as txnyr, DATEPART(MONTH, TxnDT) as txnmh, DATEPART(DAY, TxnDT) as txndy, DATEPART(Hour, TxnDT) as txnhr, SaleStartLoc    FROM Txn_RefundSpeTkt       JOIN #mTempLocation ON SaleStartLoc = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT<@Openenddate) AS T2  GROUP BY txnyr, txnmh, txndy, txnhr, SaleStartLoc 
UNION SELECT txnyr, txnmh, txndy, txnhr, 'SOneExitBack'     as mCheck, '1' AS mCheckOffset, '-' as mCheckIO, SaleEndLoc     as LocationId, COUNT(SaleEndLoc)    as num FROM (SELECT DATEPART(YEAR, TxnDT) as txnyr, DATEPART(MONTH, TxnDT) as txnmh, DATEPART(DAY, TxnDT) as txndy, DATEPART(Hour, TxnDT) as txnhr, SaleEndLoc      FROM Txn_RefundSpeTkt       JOIN #mTempLocation ON SaleEndLoc   = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT<@Openenddate) AS T2  GROUP BY txnyr, txnmh, txndy, txnhr, SaleEndLoc


DROP TABLE #mTempLocation";
        public static readonly string Sql_Command_Day_ElectronicTicket = @"--票卡進出站運量日報表
DECLARE @StartStation VARCHAR(10)	--該路線車站起點
DECLARE @EndStation VARCHAR(10)		--該路線車站終點
DECLARE @Openstardate VARCHAR(20)	--搜尋時間開始
DECLARE @Openenddate VARCHAR(20)	--搜尋時間結束
DECLARE @Monthstardate VARCHAR(20)	--開始的月份
DECLARE @Monthenddate VARCHAR(20)	--結束的月份
DECLARE @DateSet VARCHAR(20)		--計算月份用
DECLARE @RunDate VARCHAR(20)		--執行的月份

SET @StartStation = '{0}' 
SET @EndStation = '{1}'
SET @Openstardate = '{2}'
SET @Openenddate = '{3}'
SET @DateSet = CONCAT(DATEPART(YEAR,@Openstardate),'-',DATEPART(MONTH,@Openstardate) , '-01 0' , DATEPART(HOUR,@Openstardate) , ':0' , DATEPART(MINUTE,@Openstardate) , ':0' , DATEPART(SECOND,@Openstardate))
SET @Monthenddate = DATEADD(MONTH,1,@DateSet)
SET @Monthstardate = DATEADD(MONTH,-1,@Monthenddate)
SET @RunDate = @Monthstardate

SELECT
    TA.LocationId AS LocationId,
    ISNULL(TT1.Electron_Entry_num,0) - ISNULL(TT2.Student_Entry_num,0) - ISNULL(TT5.Welfare_Entry_num,0) AS Electron_Entry_num,
    ISNULL(T1.Electron_Exit_num,0) - ISNULL(T2.Student_Exit_num,0) - ISNULL(T5.Welfare_Exit_num,0) AS Electron_Exit_num,
    ISNULL(TT2.Student_Entry_num,0) AS Student_Entry_num,
    ISNULL(T2.Student_Exit_num,0) AS Student_Exit_num,
    ISNULL(TT5.Welfare_Entry_num,0) AS Welfare_Entry_num,
    ISNULL(T5.Welfare_Exit_num,0) AS Welfare_Exit_num,
    ISNULL(TT3.all_pass_Entry_num,0) - ISNULL(TT4.all_pass_Student_Entry_num,0) AS all_pass_Entry_num,
    ISNULL(T3.all_pass_Exit_num,0) - ISNULL(T4.all_pass_Student_Exit_num,0) AS all_pass_Exit_num,
    ISNULL(TT4.all_pass_Student_Entry_num,0) AS all_pass_Student_Entry_num,
    ISNULL(T4.all_pass_Student_Exit_num,0) AS all_pass_Student_Exit_num,
    0 AS MP_C_I,
    0 AS MP_C_O,
    0 AS MP_Q_I,
    0 AS MP_Q_O,
    ISNULL(TT6.SOneTkt_Entry_num - ISNULL(T11.SOneReTkt_Entry_num,0),0) AS SOneTkt_Entry_num,
    ISNULL(T6.SOneTkt_Exit_num - ISNULL(T12.SOneReTkt_Exit_num,0),0) AS SOneTkt_Exit_num,
    ISNULL(T10.SOneTkt_discount_Entry_num - ISNULL(T13.SOneReTkt_discount_Entry_num,0),0) AS SOneTkt_discount_Entry_num,
    ISNULL(TT10.SOneTkt_discount_Exit_num - ISNULL(T14.SOneReTkt_discount_Exit_num,0),0) AS SOneTkt_discount_Exit_num,
    ISNULL(TT7.SOneTkt_bike_Entry_num - ISNULL(T15.SOneReTkt_bike_Entry_num,0),0) AS SOneTkt_bike_Entry_num,
    ISNULL(T7.SOneTkt_bike_Exit_num - ISNULL(T16.SOneReTkt_bike_Exit_num,0),0) AS SOneTkt_bike_Exit_num,
    ISNULL(TT8.officeman_Entry_num,0) AS officeman_Entry_num,
    ISNULL(T8.officeman_Exit_num,0) AS officeman_Exit_num,
    ISNULL(TT9.OneDay_Entry_num,0) AS OneDay_Entry_num,
    ISNULL(T9.OneDay_Exit_num,0) AS OneDay_Exit_num
FROM 
    (SELECT LocationId  FROM Parm081_LocList WHERE CurrentType IN (1) AND LocationId BETWEEN @StartStation AND @EndStation) AS TA 
    LEFT JOIN (SELECT LocationId,     COUNT(*) AS Electron_Entry_num              FROM Txn_Entry              WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and IssueCode != 177 and ((TxnType in (21) ) or (TxnType in (23)and IssueCode in (11))) GROUP BY LocationId) AS TT1 ON TA.LocationId = TT1.LocationId 
    LEFT JOIN (SELECT LocationId,     COUNT(*) AS Electron_Exit_num               FROM Txn_Exit               WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and IssueCode != 177 and ((TxnType in (22) ) or (TxnType in (24)and IssueCode in (11))) GROUP BY LocationId) AS T1 ON TA.LocationId = T1.LocationId 
    LEFT JOIN (SELECT LocationId,     COUNT(*) AS Student_Entry_num               FROM Txn_Entry              WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and IssueCode != 177 and TxnType in(21) and((IssueCode in (2) and PersonalProfile in (5)) or(IssueCode in (9) and IdentityType in(2) and CardType in (7)AND AreaCode not IN (24,25,30,34,38,42,46,50,54,58,62,66,70,74,78,82,86,90,94,98,102,106,110,114)) or(IssueCode in (11) and CardType in (2))) and DATEDIFF(MINUTE, TxnDT, IdentityExpiryDT) > 0 GROUP BY LocationId) AS TT2 ON TA.LocationId = TT2.LocationId 
    LEFT JOIN (SELECT LocationId,     COUNT(*) AS Student_Exit_num                FROM Txn_Exit               WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and IssueCode != 177 and DATEDIFF(MINUTE, TxnDT, IdentityExpiryDT) > 0 and TxnType in(22) and ((IssueCode in (2) and PersonalProfile in (5)) or (IssueCode in (9) and IdentityType in(2) and CardType in (7)AND AreaCode not IN (24,25,30,34,38,42,46,50,54,58,62,66,70,74,78,82,86,90,94,98,102,106,110,114)) or (IssueCode in (11) and CardType in (2))) GROUP BY LocationId) AS T2 ON TA.LocationId = T2.LocationId 
    LEFT JOIN (SELECT LocationId,     COUNT(*) AS all_pass_Entry_num              FROM Txn_Entry              WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and IssueCode != 177 and TxnType in(27,29) GROUP BY LocationId) AS TT3 ON TA.LocationId = TT3.LocationId  
    LEFT JOIN (SELECT LocationId,     COUNT(*) AS all_pass_Exit_num               FROM Txn_Exit               WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and IssueCode != 177 and TxnType in(28,30) GROUP BY LocationId) AS T3 ON TA.LocationId = T3.LocationId 
    LEFT JOIN (SELECT LocationId,     COUNT(*) AS all_pass_Student_Entry_num      FROM Txn_Entry              WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and IssueCode != 177 and TxnType in(27,29) and ((IssueCode in (2) and PersonalProfile in (5)) or (IssueCode in (9) and IdentityType in(2) and CardType in (7)) or (IssueCode in (11) and CardType in (2))) GROUP BY LocationId) AS TT4 ON TA.LocationId = TT4.LocationId 
    LEFT JOIN (SELECT LocationId,     COUNT(*) AS all_pass_Student_Exit_num       FROM Txn_Exit               WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and IssueCode != 177 and TxnType in(28,30) and ((IssueCode in (2) and PersonalProfile in (5)) or (IssueCode in (9) and IdentityType in(2) and CardType in (7)) or (IssueCode in (11) and CardType in (2))) GROUP BY LocationId) AS T4 ON TA.LocationId = T4.LocationId 
    LEFT JOIN (SELECT LocationId,     COUNT(*) AS Welfare_Entry_num               FROM Txn_Entry              WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and DATEDIFF(MINUTE, TxnDT, IdentityExpiryDT) > 0  and TxnType in (21) and ((IssueCode = 2 and PersonalProfile in (1, 2, 3, 4)) or (IssueCode = 2 and PersonalProfile in (8) and AreaCode in (1, 2)) or (IssueCode = 9 and IdentityType in (3, 4, 6, 5)) or (IssueCode = 9 AND IdentityType IN (2) AND AreaCode IN (119,118)) OR (IssueCode = 9 AND IdentityType IN (2) AND AreaCode IN (24,25,30,34,38,42,46,50,54,58,62,66,70,74,78,82,86,90,94,98,102,106,110,114))) GROUP BY LocationId) AS TT5 ON TA.LocationId = TT5.LocationId  
    LEFT JOIN (SELECT LocationId,     COUNT(*) AS Welfare_Exit_num                FROM Txn_Exit               WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and DATEDIFF(MINUTE, TxnDT, IdentityExpiryDT) > 0 and TxnType in (22) and ((IssueCode = 2 and PersonalProfile in (1, 2, 3, 4)) or (IssueCode = 2 and PersonalProfile in (8) and AreaCode in (1, 2)) or (IssueCode = 9 and IdentityType in (3, 4, 6, 5)) or (IssueCode = 9 AND IdentityType IN (2) AND AreaCode IN (119,118)) OR (IssueCode = 9 AND IdentityType IN (2) AND AreaCode IN (24,25,30,34,38,42,46,50,54,58,62,66,70,74,78,82,86,90,94,98,102,106,110,114))) GROUP BY LocationId) AS T5 ON TA.LocationId = T5.LocationId 
    LEFT JOIN (SELECT StartLoc,       COUNT(*) AS SOneTkt_Entry_num               FROM Txn_SellSpecialTicket  WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and CardType in(1) GROUP BY StartLoc) AS TT6 ON TA.LocationId = TT6.StartLoc 
    LEFT JOIN (SELECT EndLoc,         COUNT(*) AS SOneTkt_Exit_num                FROM Txn_SellSpecialTicket  WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and CardType in(1) GROUP BY EndLoc) AS T6 ON TA.LocationId = T6.EndLoc 
    LEFT JOIN (SELECT StartLoc,       COUNT(*) AS SOneTkt_discount_Entry_num      FROM Txn_SellSpecialTicket  WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and CardType in(3) GROUP BY StartLoc) AS T10 ON TA.LocationId = T10.StartLoc 
    LEFT JOIN (SELECT EndLoc,         COUNT(*) AS SOneTkt_discount_Exit_num       FROM Txn_SellSpecialTicket  WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and CardType in(3) GROUP BY EndLoc) AS TT10 ON TA.LocationId = TT10.EndLoc 
    LEFT JOIN (SELECT StartLoc,       COUNT(*) AS SOneTkt_bike_Entry_num          FROM Txn_SellSpecialTicket  WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and CardType in(2) GROUP BY StartLoc) AS TT7 ON TA.LocationId = TT7.StartLoc 
    LEFT JOIN (SELECT EndLoc,         COUNT(*) AS SOneTkt_bike_Exit_num           FROM Txn_SellSpecialTicket  WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and CardType in(2) GROUP BY EndLoc) AS T7 ON TA.LocationId = T7.EndLoc 
    LEFT JOIN (SELECT SaleStartLoc,   COUNT(*) AS SOneReTkt_Entry_num             FROM Txn_RefundSpeTkt       WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and CardType in(1) GROUP BY SaleStartLoc) AS T11 ON TA.LocationId = T11.SaleStartLoc 
    LEFT JOIN (SELECT SaleEndLoc,     COUNT(*) AS SOneReTkt_Exit_num              FROM Txn_RefundSpeTkt       WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and CardType in(1) GROUP BY SaleEndLoc) AS T12 ON TA.LocationId = T12.SaleEndLoc 
    LEFT JOIN (SELECT SaleStartLoc,   COUNT(*) AS SOneReTkt_discount_Entry_num    FROM Txn_RefundSpeTkt       WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and CardType in(3) GROUP BY SaleStartLoc) AS T13 ON TA.LocationId = T13.SaleStartLoc 
    LEFT JOIN (SELECT SaleEndLoc,     COUNT(*) AS SOneReTkt_discount_Exit_num     FROM Txn_RefundSpeTkt       WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and CardType in(3) GROUP BY SaleEndLoc) AS T14 ON TA.LocationId = T14.SaleEndLoc 
    LEFT JOIN (SELECT SaleStartLoc,   COUNT(*) AS SOneReTkt_bike_Entry_num        FROM Txn_RefundSpeTkt       WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and CardType in(2) GROUP BY SaleStartLoc) AS T15 ON TA.LocationId = T15.SaleStartLoc 
    LEFT JOIN (SELECT SaleEndLoc,     COUNT(*) AS SOneReTkt_bike_Exit_num         FROM Txn_RefundSpeTkt       WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and CardType in(2) GROUP BY SaleEndLoc) AS T16 ON TA.LocationId = T16.SaleEndLoc 
    LEFT JOIN (SELECT LocationId,     COUNT(*) AS officeman_Entry_num             FROM Txn_Entry              WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and TxnType = 23 AND PeriodCode = 17 GROUP BY LocationId) AS TT8 ON TA.LocationId = TT8.LocationId 
    LEFT JOIN (SELECT LocationId,     COUNT(*) AS officeman_Exit_num              FROM Txn_Exit               WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and TxnType = 24 AND PeriodCode = 17 GROUP BY LocationId) AS T8 ON TA.LocationId = T8.LocationId
    LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS OneDay_Entry_num      FROM Txn_Entry              WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate and TxnType = 23 and (IssueCode = 9 and PeriodCode = 1) GROUP BY LocationId) AS TT9 ON TA.LocationId = TT9.LocationId
    LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS OneDay_Exit_num       FROM Txn_Exit               WHERE LocationId >= @StartStation and LocationId <= @EndStation and  TxnDT >= @Openstardate and TxnDT < @Openenddate  and TxnType = 24 and (IssueCode = 9 and PeriodCode = 1) GROUP BY LocationId) AS T9 ON TA.LocationId = T9.LocationId 
ORDER BY LocationId";
        public static readonly string Sql_Command_Day_AllRideList = @"";
        public static readonly string Sql_Command_Day_TrafficAmount = @"--營收日報表
DECLARE @StartStation VARCHAR(10)	--該路線車站起點
DECLARE @EndStation VARCHAR(10)		--該路線車站終點
DECLARE @Openstardate VARCHAR(20)	--搜尋時間開始
DECLARE @Openenddate VARCHAR(20)	--搜尋時間結束
DECLARE @Monthstardate VARCHAR(20)	--開始的月份
DECLARE @Monthenddate VARCHAR(20)	--結束的月份
DECLARE @DateSet VARCHAR(20)		--計算月份用
DECLARE @RunDate VARCHAR(20)		--執行的月份

SET @StartStation = '{0}' 
SET @EndStation = '{1}'
SET @Openstardate = '{2}'
SET @Openenddate = '{3}'
SET @DateSet = CONCAT(DATEPART(YEAR,@Openstardate),'-',DATEPART(MONTH,@Openstardate) , '-01 0' , DATEPART(HOUR,@Openstardate) , ':0' , DATEPART(MINUTE,@Openstardate) , ':0' , DATEPART(SECOND,@Openstardate))
SET @Monthenddate = DATEADD(MONTH,1,@DateSet)
SET @Monthstardate = DATEADD(MONTH,-1,@Monthenddate)
SET @RunDate = @Monthstardate

SELECT 
    TA.LocationId,
    (ISNULL(T1.ECCTxnAmt,0) + ISNULL(T8.TxnAmtCsc,0) + ISNULL(T9.TxnAmtDed,0)) AS ECCTxnAmt,
    ISNULL(T2.IPASSTxnAmt,0) AS IPASSTxnAmt,
    ISNULL(T3.ICASHTxnAmt,0) AS ICASHTxnAmt,

    0 AS CreditCard,
    0 AS QrCode,

    ISNULL(T4.SOneTkt_Amt,0) - ISNULL(T10.TxnAmtRe,0)AS SOneTkt_Amt,
    ISNULL(T5.SOneTkt_Amt_discount,0) AS SOneTkt_Amt_discount,
    ISNULL(T7.SOneTkt_bike_Amt,0) AS SOneTkt_bike_Amt,
    0 AS OneDayTicket,
    0 AS OneDayTicket_Law,
    0 AS GroupTicket,
    0 AS OpenTrafficSetTicket,
    0 AS OwnTicketSupply,
    0 AS VavmTackOut,
    0 AS AddValeFail,
    0 AS RefundCard,
    0 AS RefundPv
    
FROM ( 
               SELECT LocationId  FROM Parm081_LocList WHERE CurrentType IN (1) AND LocationId BETWEEN @StartStation AND @EndStation) AS TA 
    LEFT JOIN (SELECT LocationId, ISNULL(SUM(TxnAmt),0) AS ECCTxnAmt               FROM Txn_Exit               WHERE TxnDT >= @Openstardate and TxnDT < @Openenddate and TxnType in(22) and IssueCode = 2  GROUP BY LocationId ) AS T1 ON TA.LocationId = T1.LocationId 
    LEFT JOIN (SELECT LocationId, ISNULL(SUM(TxnAmt),0) AS IPASSTxnAmt             FROM Txn_Exit               WHERE TxnDT >= @Openstardate and TxnDT < @Openenddate and TxnType in(22) and IssueCode = 9  GROUP BY LocationId ) AS T2 ON TA.LocationId = T2.LocationId 
    LEFT JOIN (SELECT LocationId, ISNULL(SUM(TxnAmt),0) AS ICASHTxnAmt             FROM Txn_Exit               WHERE TxnDT >= @Openstardate and TxnDT < @Openenddate and TxnType in(22) and IssueCode = 11 GROUP BY LocationId ) AS T3 ON TA.LocationId = T3.LocationId 
    LEFT JOIN (SELECT LocationId, ISNULL(SUM(TxnAmt),0) AS SOneTkt_Amt             FROM Txn_SellSpecialTicket  WHERE TxnDT >= @Openstardate and TxnDT < @Openenddate and CardType in(1)                    GROUP BY LocationId ) AS T4 ON TA.LocationId = T4.LocationId 
    LEFT JOIN (SELECT LocationId, ISNULL(SUM(TxnAmt),0) AS SOneTkt_Amt_discount    FROM Txn_SellSpecialTicket  WHERE TxnDT >= @Openstardate and TxnDT < @Openenddate and CardType in(3)                    GROUP BY LocationId ) AS T5 ON TA.LocationId = T5.LocationId 
    LEFT JOIN (SELECT LocationId, ISNULL(SUM(TxnAmt),0) AS SOneTkt_bike_Amt        FROM Txn_SellSpecialTicket  WHERE TxnDT >= @Openstardate and TxnDT < @Openenddate and CardType in(2)                    GROUP BY LocationId ) AS T7 ON TA.LocationId = T7.LocationId 
    LEFT JOIN (SELECT LocationId, ISNULL(SUM(TxnAmt),0) AS TxnAmtCsc               FROM Txn_ExcessByCSC        WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate                                       GROUP BY LocationId ) AS T8 ON TA.LocationId = T8.LocationId 
    LEFT JOIN (SELECT LocationId, ISNULL(SUM(TxnAmt),0) AS TxnAmtDed               FROM Txn_ExceptDeduct       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate                                       GROUP BY LocationId ) AS T9 ON TA.LocationId = T9.LocationId 
    LEFT JOIN (SELECT LocationId, ISNULL(SUM(TxnAmt),0) AS TxnAmtRe                FROM Txn_RefundSpeTkt       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (97,98,99)             GROUP BY LocationId ) AS T10 ON TA.LocationId = T10.LocationId 
ORDER BY LocationId
";
        public static readonly string Sql_Command_Day_OriginDestination = @"--每日起迄總表
DECLARE @StartStation VARCHAR(10)	--該路線車站起點
DECLARE @EndStation VARCHAR(10)		--該路線車站終點
DECLARE @Openstardate VARCHAR(20)	--搜尋時間開始
DECLARE @Openenddate VARCHAR(20)	--搜尋時間結束
DECLARE @Monthstardate VARCHAR(20)	--開始的月份
DECLARE @Monthenddate VARCHAR(20)	--結束的月份
DECLARE @DateSet VARCHAR(20)		--計算月份用
DECLARE @RunDate VARCHAR(20)		--執行的月份
DECLARE @in_columns VARCHAR(100)	--存放車站別

SET @StartStation = '{0}' 
SET @EndStation = '{1}'
SET @Openstardate = '{2}'
SET @Openenddate = '{3}'
SET @DateSet = CONCAT(DATEPART(YEAR,@Openstardate),'-',DATEPART(MONTH,@Openstardate) , '-01 0' , DATEPART(HOUR,@Openstardate) , ':0' , DATEPART(MINUTE,@Openstardate) , ':0' , DATEPART(SECOND,@Openstardate))
SET @Monthenddate = DATEADD(MONTH,1,@DateSet)
SET @Monthstardate = DATEADD(MONTH,-1,@Monthenddate)
SET @RunDate = @Monthstardate

SELECT LocationId as mLoc INTO #TempLocationId FROM Parm081_LocList WHERE CurrentType IN (1) AND LocationId BETWEEN @StartStation AND @EndStation ORDER BY LocationId

select @in_columns = COALESCE(@in_columns + ',[' + cast(mLoc as VARCHAR) + ']','[' + cast(mLoc as VARCHAR) + ']') from #TempLocationId --GROUP BY QUOTENAME(mLoc) 

--set @in_columns  = '[101],[102]'

SELECT StartLoc, EndLoc, ISNULL(SUM(COUNTs),0) AS COUNTs, '+' AS mCheckIO 
into #tempList
FROM (
          SELECT EntryLocation AS StartLoc, LocationId AS EndLoc,   ISNULL(COUNT(*),0) AS COUNTs    FROM Txn_Exit               WHERE @Openstardate < TxnDT AND TxnDT < @Openenddate                GROUP BY EntryLocation, LocationId 
    UNION SELECT StartLoc,                EndLoc,                   ISNULL(COUNT(*),0) AS COUNTs    FROM Txn_SellSpecialTicket  WHERE @Openstardate < TxnDT AND TxnDT < @Openenddate                GROUP BY StartLoc,EndLoc 
    UNION SELECT EntryLocation AS StartLoc, ExitLocation AS EndLoc, ISNULL(COUNT(*),0) AS COUNTs    FROM Txn_ExceptDeduct       WHERE @Openstardate < TxnDT AND TxnDT < @Openenddate AND TxnAmt > 0 GROUP BY EntryLocation,ExitLocation 
) AS T1 JOIN #TempLocationId AS TL ON T1.StartLoc = TL.mLoc
GROUP BY T1.StartLoc,T1.EndLoc 
UNION SELECT SaleStartLoc, SaleEndLoc, 0 - ISNULL(COUNT(*),0) AS COUNTs, '-' AS mCheckIO FROM Txn_RefundSpeTkt WHERE @Openstardate < TxnDT AND TxnDT < @Openenddate GROUP BY SaleStartLoc, SaleEndLoc
ORDER BY StartLoc, EndLoc

EXECUTE('
select *
from (select EndLoc, StartLoc, COUNTs from #tempList) as t
pivot (sum(COUNTs) for StartLoc in ('+ @in_columns +')) as p

');

DROP TABLE #tempList
DROP TABLE #TempLocationId
";
        public static readonly string Sql_Command_Day_OriginDestination_Old = @"--每日起迄總表
DECLARE @StartStation VARCHAR(10)	--該路線車站起點
DECLARE @EndStation VARCHAR(10)		--該路線車站終點
DECLARE @Openstardate VARCHAR(20)	--搜尋時間開始
DECLARE @Openenddate VARCHAR(20)	--搜尋時間結束
DECLARE @Monthstardate VARCHAR(20)	--開始的月份
DECLARE @Monthenddate VARCHAR(20)	--結束的月份
DECLARE @DateSet VARCHAR(20)		--計算月份用
DECLARE @RunDate VARCHAR(20)		--執行的月份
DECLARE @in_columns VARCHAR(100)	--存放車站別

SET @StartStation = '{0}' 
SET @EndStation = '{1}'
SET @Openstardate = '{2}'
SET @Openenddate = '{3}'
SET @DateSet = CONCAT(DATEPART(YEAR,@Openstardate),'-',DATEPART(MONTH,@Openstardate) , '-01 0' , DATEPART(HOUR,@Openstardate) , ':0' , DATEPART(MINUTE,@Openstardate) , ':0' , DATEPART(SECOND,@Openstardate))
SET @Monthenddate = DATEADD(MONTH,1,@DateSet)
SET @Monthstardate = DATEADD(MONTH,-1,@Monthenddate)
SET @RunDate = @Monthstardate

SELECT LocationId as mLoc INTO #TempLocationId FROM Parm081_LocList WHERE CurrentType IN (1) AND LocationId BETWEEN @StartStation AND @EndStation ORDER BY LocationId

select @in_columns = COALESCE(@in_columns + ',[' + cast(mLoc as VARCHAR) + ']','[' + cast(mLoc as VARCHAR) + ']') from #TempLocationId --GROUP BY QUOTENAME(mLoc) 

--set @in_columns  = '[101],[102]'

SELECT StartLoc, EndLoc, ISNULL(SUM(COUNTs),0) AS COUNTs, '+' AS mCheckIO 
into #tempList
FROM (
          SELECT EntryLocation AS StartLoc, LocationId AS EndLoc,   ISNULL(COUNT(*),0) AS COUNTs    FROM Txn_Exit               WHERE @Openstardate < TxnDT AND TxnDT < @Openenddate                GROUP BY EntryLocation, LocationId 
    UNION SELECT StartLoc,                EndLoc,                   ISNULL(COUNT(*),0) AS COUNTs    FROM Txn_SellSpecialTicket  WHERE @Openstardate < TxnDT AND TxnDT < @Openenddate                GROUP BY StartLoc,EndLoc 
    UNION SELECT EntryLocation AS StartLoc, ExitLocation AS EndLoc, ISNULL(COUNT(*),0) AS COUNTs    FROM Txn_ExceptDeduct       WHERE @Openstardate < TxnDT AND TxnDT < @Openenddate AND TxnAmt > 0 GROUP BY EntryLocation,ExitLocation 
) AS T1 JOIN #TempLocationId AS TL ON T1.StartLoc = TL.mLoc
GROUP BY T1.StartLoc,T1.EndLoc 
UNION SELECT SaleStartLoc, SaleEndLoc, 0 - ISNULL(COUNT(*),0) AS COUNTs, '-' AS mCheckIO FROM Txn_RefundSpeTkt WHERE @Openstardate < TxnDT AND TxnDT < @Openenddate GROUP BY SaleStartLoc, SaleEndLoc
ORDER BY StartLoc, EndLoc

EXECUTE('
select *
from (select EndLoc, StartLoc, COUNTs from #tempList) as t
pivot (sum(COUNTs) for StartLoc in ('+ @in_columns +')) as p

');

DROP TABLE #tempList
DROP TABLE #TempLocationId
";
        public static readonly string Sql_Command_Day_EquipAmount = @"--設備營收日報表
DECLARE @StartStation VARCHAR(10)	--該路線車站起點
DECLARE @EndStation VARCHAR(10)		--該路線車站終點
DECLARE @Openstardate VARCHAR(20)	--搜尋時間開始
DECLARE @Openenddate VARCHAR(20)	--搜尋時間結束
DECLARE @Monthstardate VARCHAR(20)	--開始的月份
DECLARE @Monthenddate VARCHAR(20)	--結束的月份
DECLARE @DateSet VARCHAR(20)		--計算月份用
DECLARE @RunDate VARCHAR(20)		--執行的月份

SET @StartStation = '{0}' 
SET @EndStation = '{1}'
SET @Openstardate = '{2}'
SET @Openenddate = '{3}'
SET @DateSet = CONCAT(DATEPART(YEAR,@Openstardate),'-',DATEPART(MONTH,@Openstardate) , '-01 0' , DATEPART(HOUR,@Openstardate) , ':0' , DATEPART(MINUTE,@Openstardate) , ':0' , DATEPART(SECOND,@Openstardate))
SET @Monthenddate = DATEADD(MONTH,1,@DateSet)
SET @Monthstardate = DATEADD(MONTH,-1,@Monthenddate)
SET @RunDate = @Monthstardate

SELECT 
    TA.LocationId, 
    ISNULL(E1.ECC_TxnAmt, 0) AS ECC_TxnAmt, 
    ISNULL(E2.ECC_CSCAmt, 0) AS ECC_CSCAmt, 
    ISNULL(E9.ECC_ExceptDeductAmt, 0) AS ECC_ExceptDeductAmt, 
    ISNULL(E3.ECC_SaleAmt, 0) AS ECC_SaleAmt, 
    ISNULL(ISNULL(E4.ECC_VAVMAddValueAmt, 0) - ISNULL(E5.ECC_VAVMCancelValueAmt, 0), 0) AS ECC_AddValueAmt, 
    ISNULL(E6.ECC_PAMAddValueAmt, 0) AS ECC_PAMAddValueAmt, 
    ISNULL(E7.ECC_PAMCancelValueAmt, 0) AS ECC_PAMCancelValueAmt, 
    ISNULL(E8.ECC_AutoAddValueAmt, 0) AS ECC_AutoAddValueAmt,

    ISNULL(P1.IPASS_TxnAmt, 0) AS IPASS_TxnAmt, 
    ISNULL(P2.IPASS_CSCAmt, 0) AS IPASS_CSCAmt, 
    ISNULL(P9.IPASS_ExceptDeductAmt, 0) AS IPASS_ExceptDeductAmt, 
    ISNULL(P3.IPASS_SaleAmt, 0) AS IPASS_SaleAmt, 
    ISNULL(ISNULL(P4.IPASS_VAVMAddValueAmt, 0) - ISNULL(P5.IPASS_VAVMCancelValueAmt, 0), 0) AS IPASS_AddValueAmt, 
    ISNULL(P6.IPASS_PAMAddValueAmt, 0) AS IPASS_PAMAddValueAmt, 
    ISNULL(P7.IPASS_PAMCancelValueAmt, 0) AS IPASS_PAMCancelValueAmt, 
    ISNULL(P8.IPASS_AutoAddValueAmt, 0) AS IPASS_AutoAddValueAmt,

    ISNULL(C1.ICASH_TxnAmt, 0) AS ICASH_TxnAmt, 
    ISNULL(C2.ICASH_CSCAmt, 0) AS ICASH_CSCAmt, 
    ISNULL(C9.ICASH_ExceptDeductAmt, 0) AS ICASH_ExceptDeductAmt, 
    ISNULL(C3.ICASH_SaleAmt, 0) AS ICASH_SaleAmt, 
    ISNULL(ISNULL(C4.ICASH_VAVMAddValueAmt, 0) - ISNULL(C5.ICASH_VAVMCancelValueAmt, 0), 0) AS ICASH_AddValueAmt, 
    ISNULL(C6.ICASH_PAMAddValueAmt, 0) AS ICASH_PAMAddValueAmt, 
    ISNULL(C7.ICASH_PAMCancelValueAmt, 0) AS ICASH_PAMCancelValueAmt, 
    ISNULL(C8.ICASH_AutoAddValueAmt, 0) AS ICASH_AutoAddValueAmt,

    0 AS MobilePay_MasterCard_Exit,
    0 AS MobilePay_MasterCard_Fare,
    0 AS MobilePay_CUP_Exit,
    0 AS MobilePay_CUP_Fare,
    0 AS MobilePay_Discover_Exit,
    0 AS MobilePay_Discover_Fare,
    0 AS MobilePay_LinePayMoney_Exit,
    0 AS MobilePay_LinePayMoney_Fare,
    0 AS MobilePay_TSQrCode_Exit,
    0 AS MobilePay_TSQrCode_Fare,

    ISNULL(T1.SOneTkt_Amt, 0) AS SOneTkt_Amt, 
    ISNULL(T2.SOneTkt_Amt_discount, 0) AS SOneTkt_Amt_discount, 
    ISNULL(T3.SOneTkt_bike_Amt, 0) AS SOneTkt_bike_Amt, 
    ISNULL(T4.PTADuty, 0) AS PTADuty, 
    ISNULL(T5.RefundSpeTkt, 0) AS RefundSpeTkt
FROM(
              SELECT LocationId  FROM Parm081_LocList WHERE CurrentType IN (1) AND LocationId BETWEEN @StartStation AND @EndStation) AS TA
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ECC_TxnAmt               FROM Txn_Exit              WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (22) AND IssueCode = 2                         GROUP BY LocationId) AS E1 ON TA.LocationId = E1.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ECC_CSCAmt               FROM Txn_ExcessByCSC       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (71) AND IssueCode = 2                         GROUP BY LocationId) AS E2 ON TA.LocationId = E2.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ECC_ExceptDeductAmt      FROM Txn_ExceptDeduct      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND IssueCode = 2 AND TxnType = 33                            GROUP BY LocationId) AS E9 ON TA.LocationId = E9.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ECC_SaleAmt              FROM Txn_SaleCard          WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (11) AND IssueCode = 2                         GROUP BY LocationId) AS E3 ON TA.LocationId = E3.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ECC_VAVMAddValueAmt      FROM Txn_AddValue          WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (1) AND IssueCode = 2 AND EquipType = 5        GROUP BY LocationId) AS E4 ON TA.LocationId = E4.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ECC_VAVMCancelValueAmt   FROM Txn_CancelValue       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (5) AND IssueCode = 2 AND EquipType = 5        GROUP BY LocationId) AS E5 ON TA.LocationId = E5.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ECC_PAMAddValueAmt       FROM Txn_AddValue          WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (1) AND IssueCode = 2 AND EquipType IN(3)      GROUP BY LocationId) AS E6 ON TA.LocationId = E6.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ECC_PAMCancelValueAmt    FROM Txn_CancelValue       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (5) AND IssueCode = 2 AND EquipType IN(3)      GROUP BY LocationId) AS E7 ON TA.LocationId = E7.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ECC_AutoAddValueAmt      FROM ECCTxn_Interface      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType = 2 AND MyTxnType = 2                             GROUP BY LocationId) AS E8 ON TA.LocationId = E8.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS IPASS_TxnAmt             FROM Txn_Exit              WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (22) AND IssueCode = 9                         GROUP BY LocationId) AS P1 ON TA.LocationId = P1.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS IPASS_CSCAmt             FROM Txn_ExcessByCSC       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (71) AND IssueCode = 9                         GROUP BY LocationId) AS P2 ON TA.LocationId = P2.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS IPASS_ExceptDeductAmt    FROM Txn_ExceptDeduct      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND IssueCode = 9 AND TxnType = 33                            GROUP BY LocationId) AS P9 ON TA.LocationId = P9.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS IPASS_SaleAmt            FROM Txn_SaleCard          WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (11) AND IssueCode = 9                         GROUP BY LocationId) AS P3 ON TA.LocationId = P3.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS IPASS_VAVMAddValueAmt    FROM Txn_AddValue          WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (1) AND IssueCode = 9 AND EquipType = 5        GROUP BY LocationId) AS P4 ON TA.LocationId = P4.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS IPASS_VAVMCancelValueAmt FROM Txn_CancelValue       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (1) AND IssueCode = 9 AND EquipType = 5        GROUP BY LocationId) AS P5 ON TA.LocationId = P5.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS IPASS_PAMAddValueAmt     FROM Txn_AddValue          WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (1) AND IssueCode = 9 AND EquipType IN(3)      GROUP BY LocationId) AS P6 ON TA.LocationId = P6.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS IPASS_PAMCancelValueAmt  FROM Txn_CancelValue       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (5) AND IssueCode = 9 AND EquipType IN(3)      GROUP BY LocationId) AS P7 ON TA.LocationId = P7.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS IPASS_AutoAddValueAmt    FROM ECCTxn_Interface      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType = 66                                              GROUP BY LocationId) AS P8 ON TA.LocationId = P8.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ICASH_TxnAmt             FROM Txn_Exit              WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (22) AND IssueCode = 11                        GROUP BY LocationId) AS C1 ON TA.LocationId = C1.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ICASH_CSCAmt             FROM Txn_ExcessByCSC       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (71) AND IssueCode = 11                        GROUP BY LocationId) AS C2 ON TA.LocationId = C2.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ICASH_ExceptDeductAmt    FROM Txn_ExceptDeduct      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND IssueCode = 11 AND TxnType = 33                           GROUP BY LocationId) AS C9 ON TA.LocationId = C9.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ICASH_SaleAmt            FROM Txn_SaleCard          WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (11) AND IssueCode = 11                        GROUP BY LocationId) AS C3 ON TA.LocationId = C3.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ICASH_VAVMAddValueAmt    FROM Txn_AddValue          WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (1) AND IssueCode = 11 AND EquipType = 5       GROUP BY LocationId) AS C4 ON TA.LocationId = C4.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ICASH_VAVMCancelValueAmt FROM Txn_CancelValue       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (5) AND IssueCode = 11 AND EquipType = 5       GROUP BY LocationId) AS C5 ON TA.LocationId = C5.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ICASH_PAMAddValueAmt     FROM Txn_AddValue          WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (1) AND IssueCode = 11 AND EquipType IN(3)     GROUP BY LocationId) AS C6 ON TA.LocationId = C6.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ICASH_PAMCancelValueAmt  FROM Txn_CancelValue       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (5) AND IssueCode = 11 AND EquipType IN(3)     GROUP BY LocationId) AS C7 ON TA.LocationId = C7.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ICASH_AutoAddValueAmt    FROM ECCTxn_Interface      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType = 2 AND TxnSubType = 4                            GROUP BY LocationId) AS C8 ON TA.LocationId = C8.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS SOneTkt_Amt              FROM Txn_SellSpecialTicket WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND CardType in (1) AND EquipType Not IN (6)                  GROUP BY LocationId) AS T1 ON TA.LocationId = T1.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS SOneTkt_Amt_discount     FROM Txn_SellSpecialTicket WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND CardType in (3) AND EquipType Not IN (6)                  GROUP BY LocationId) AS T2 ON TA.LocationId = T2.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS SOneTkt_bike_Amt         FROM Txn_SellSpecialTicket WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND CardType in (2) AND EquipType Not IN (6)                  GROUP BY LocationId) AS T3 ON TA.LocationId = T3.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS PTADuty                  FROM Txn_SellSpecialTicket WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND EquipType IN (6)                                          GROUP BY LocationId) AS T4 ON TA.LocationId = T4.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS RefundSpeTkt             FROM Txn_RefundSpeTkt      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate                                                               GROUP BY LocationId) AS T5 ON TA.LocationId = T5.LocationId
ORDER BY LocationId 
";
        public static readonly string Sql_Command_Month_AllRideList = @"";
        public static readonly string Sql_Command_Month_OriginDestination = @"--每月起迄總表
DECLARE @StartStation VARCHAR(10)	--該路線車站起點
DECLARE @EndStation VARCHAR(10)		--該路線車站終點
DECLARE @Openstardate VARCHAR(20)	--搜尋時間開始
DECLARE @Openenddate VARCHAR(20)	--搜尋時間結束
DECLARE @Monthstardate VARCHAR(20)	--開始的月份
DECLARE @Monthenddate VARCHAR(20)	--結束的月份
DECLARE @DateSet VARCHAR(20)		--計算月份用
DECLARE @RunDate VARCHAR(20)		--執行的月份
DECLARE @in_columns VARCHAR(100)	--存放車站別

SET @StartStation = '{0}' 
SET @EndStation = '{1}'
SET @Openstardate = '{2}'
SET @Openenddate = '{3}'
SET @DateSet = CONCAT(DATEPART(YEAR,@Openstardate),'-',DATEPART(MONTH,@Openstardate) , '-01 0' , DATEPART(HOUR,@Openstardate) , ':0' , DATEPART(MINUTE,@Openstardate) , ':0' , DATEPART(SECOND,@Openstardate))
SET @Monthenddate = DATEADD(MONTH,1,@DateSet)
SET @Monthstardate = DATEADD(MONTH,-1,@Monthenddate)
SET @RunDate = @Monthstardate

SELECT LocationId as mLoc INTO #TempLocationId FROM Parm081_LocList WHERE CurrentType IN (1) AND LocationId BETWEEN @StartStation AND @EndStation ORDER BY LocationId

select @in_columns = COALESCE(@in_columns + ',[' + cast(mLoc as VARCHAR) + ']','[' + cast(mLoc as VARCHAR) + ']') from #TempLocationId --GROUP BY QUOTENAME(mLoc) 

--set @in_columns  = '[101],[102]'

SELECT StartLoc, EndLoc, ISNULL(SUM(COUNTs),0) AS COUNTs, '+' AS mCheckIO 
into #tempList
FROM (
          SELECT EntryLocation AS StartLoc, LocationId AS EndLoc,   ISNULL(COUNT(*),0) AS COUNTs    FROM Txn_Exit               WHERE @Openstardate < TxnDT AND TxnDT < @Openenddate                GROUP BY EntryLocation, LocationId 
    UNION SELECT StartLoc,                EndLoc,                   ISNULL(COUNT(*),0) AS COUNTs    FROM Txn_SellSpecialTicket  WHERE @Openstardate < TxnDT AND TxnDT < @Openenddate                GROUP BY StartLoc,EndLoc 
    UNION SELECT EntryLocation AS StartLoc, ExitLocation AS EndLoc, ISNULL(COUNT(*),0) AS COUNTs    FROM Txn_ExceptDeduct       WHERE @Openstardate < TxnDT AND TxnDT < @Openenddate AND TxnAmt > 0 GROUP BY EntryLocation,ExitLocation 
) AS T1 JOIN #TempLocationId AS TL ON T1.StartLoc = TL.mLoc
GROUP BY T1.StartLoc,T1.EndLoc 
UNION SELECT SaleStartLoc, SaleEndLoc, 0 - ISNULL(COUNT(*),0) AS COUNTs, '-' AS mCheckIO FROM Txn_RefundSpeTkt WHERE @Openstardate < TxnDT AND TxnDT < @Openenddate GROUP BY SaleStartLoc, SaleEndLoc
ORDER BY StartLoc, EndLoc

EXECUTE('
select *
from (select EndLoc, StartLoc, COUNTs from #tempList) as t
pivot (sum(COUNTs) for StartLoc in ('+ @in_columns +')) as p

');

DROP TABLE #tempList
DROP TABLE #TempLocationId
";
        public static readonly string Sql_Command_Month_ElectronicTicket_Station = @"--各家票卡運量月報表(車站)
DECLARE @StartStation VARCHAR(10)	--該路線車站起點
DECLARE @EndStation VARCHAR(10)		--該路線車站終點
DECLARE @Openstardate VARCHAR(20)	--搜尋時間開始
DECLARE @Openenddate VARCHAR(20)	--搜尋時間結束
DECLARE @Monthstardate VARCHAR(20)	--開始的月份
DECLARE @Monthenddate VARCHAR(20)	--結束的月份
DECLARE @DateSet VARCHAR(20)		--計算月份用
DECLARE @RunDate VARCHAR(20)		--執行的月份

SET @StartStation = '{0}' 
SET @EndStation = '{1}'
SET @Openstardate = '{2}'
SET @Openenddate = '{3}'
SET @DateSet = CONCAT(DATEPART(YEAR,@Openstardate),'-',DATEPART(MONTH,@Openstardate) , '-01 0' , DATEPART(HOUR,@Openstardate) , ':0' , DATEPART(MINUTE,@Openstardate) , ':0' , DATEPART(SECOND,@Openstardate))
SET @Monthenddate = DATEADD(MONTH,1,@DateSet)
SET @Monthstardate = DATEADD(MONTH,-1,@Monthenddate)
SET @RunDate = @Monthstardate

SELECT LocationId as mLoc
INTO #mTempLocation
FROM Parm081_LocList
WHERE LocationId >= @StartStation AND LocationId <= @EndStation AND CurrentType IN (1)

SELECT *
INTO #mTempEcc
FROM Txn_Exit JOIN #mTempLocation on LocationId = mLoc
WHERE TxnDT >= @Monthstardate AND TxnDT < @Monthenddate AND IssueCode IN (2) 

SELECT *
INTO #mTempIpass
FROM Txn_Exit JOIN #mTempLocation on LocationId = mLoc
WHERE TxnDT >= @Monthstardate AND TxnDT < @Monthenddate AND IssueCode IN (9) 

SELECT *
INTO #mTempIcash
FROM Txn_Exit JOIN #mTempLocation on LocationId = mLoc
WHERE TxnDT >= @Monthstardate AND TxnDT < @Monthenddate AND IssueCode IN (11) 

    SELECT 
        mLoc as LocationId, 
    --ECC
        (ISNULL(ECC_ALL,0) - ISNULL(ECC_STUDENT,0) - ISNULL(ECC_ELDERLY,0) - ISNULL(ECC_DISCARD,0) - ISNULL(ECC_DISPARCARD,0) - ISNULL(ECC_CHICARD,0) - ISNULL(ECC_ONEDAY,0) - ISNULL(ECC_OFFICEMAN,0) - ISNULL(ECC_COMMUTERELE_1,0)) AS Ecc_ElectronicCard, 
        ISNULL(ECC_STUDENT,0)         AS Ecc_StudentCard, 
        ISNULL(ECC_ELDERLY,0)         AS Ecc_ElderlyCard, 
        ISNULL(ECC_DISCARD,0)         AS Ecc_DisabledCard, 
        ISNULL(ECC_DISPARCARD,0)      AS Ecc_DisabledPartherCard, 
        ISNULL(ECC_CHICARD,0)         AS Ecc_ChildCard, 
        (ISNULL(ECC_COMMUTERELE_1,0) - ISNULL(ECC_COMMUTERSTU_1,0)) AS Ecc_Commuter_ElectronicAll_1, 
        ISNULL(ECC_COMMUTERSTU_1,0)     AS Ecc_Commuter_StudentCard_1, 
    --IPASS
		--Ipass_ALL as Ipass_ALL, 
        (ISNULL(Ipass_ALL,0) - ISNULL(Ipass_STUDENT,0) - ISNULL(Ipass_ELDERLY,0) - ISNULL(Ipass_DISCARD,0) - ISNULL(Ipass_DISPARCARD,0) - ISNULL(Ipass_CHICARD,0) - ISNULL(Ipass_ONEDAY,0) - ISNULL(Ipass_OFFICEMAN,0)) AS Ipass_ElectronicCard, 
        ISNULL(Ipass_STUDENT,0)       AS Ipass_StudentCard, 
        ISNULL(Ipass_ELDERLY,0)       AS Ipass_ElderlyCard, 
        ISNULL(Ipass_DISCARD,0)       AS Ipass_DisabledCard, 
        ISNULL(Ipass_DISPARCARD,0)    AS Ipass_DisabledPartherCard, 
        ISNULL(Ipass_CHICARD,0)       AS Ipass_ChildCard, 
    --ICASH
        (ISNULL(Icash_ALL,0) - ISNULL(Icash_STUDENT,0) - ISNULL(Icash_ONEDAY,0) - ISNULL(Icash_OFFICEMAN,0)) AS Icash_ElectronicCard, 
        ISNULL(Icash_STUDENT,0)       AS Icash_WelfareCard, 
    --Cridet
        0 AS Credit_M_Exit, 
        0 AS Credit_I_Exit, 
        0 AS Credit_D_Exit, 
        0 AS QrCode_P_Exit, 
        0 AS QrCode_T_Exit
    FROM 
                   (SELECT mLoc FROM #mTempLocation) AS TA 
    --ECC
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS ECC_ALL              FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate                                                                                                  GROUP BY LocationId)   AS ECC_ALL              ON TA.mLoc = ECC_ALL.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS ECC_STUDENT          FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND PersonalProfile IN (5)   GROUP BY LocationId)   AS ECC_STUDENT          ON TA.mLoc = ECC_STUDENT.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS ECC_ELDERLY          FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND PersonalProfile IN (1,2) GROUP BY LocationId)   AS ECC_ELDERLY          ON TA.mLoc = ECC_ELDERLY.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS ECC_DISCARD          FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND PersonalProfile IN (3)   GROUP BY LocationId)   AS ECC_DISCARD          ON TA.mLoc = ECC_DISCARD.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS ECC_DISPARCARD       FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND PersonalProfile IN (4)   GROUP BY LocationId)   AS ECC_DISPARCARD       ON TA.mLoc = ECC_DISPARCARD.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS ECC_CHICARD          FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND PersonalProfile IN (8) AND AreaCode IN (1,2)   GROUP BY LocationId)   AS ECC_CHICARD          ON TA.mLoc = ECC_CHICARD.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS ECC_ONEDAY           FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (24) AND PeriodCode IN (1)                                                        GROUP BY LocationId)   AS ECC_ONEDAY           ON TA.mLoc = ECC_ONEDAY.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS ECC_TIMESCARD        FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (24) AND PeriodCode IN (16)                                                       GROUP BY LocationId)   AS ECC_TIMESCARD        ON TA.mLoc = ECC_TIMESCARD.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS ECC_OFFICEMAN        FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (24) AND PeriodCode IN (17)                                                       GROUP BY LocationId)   AS ECC_OFFICEMAN        ON TA.mLoc = ECC_OFFICEMAN.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS ECC_COMMUTERELE_1    FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (28,30)                                                                              GROUP BY LocationId)   AS ECC_COMMUTERELE      ON TA.mLoc = ECC_COMMUTERELE.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS ECC_COMMUTERSTU_1    FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (28,30) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND PersonalProfile IN (5)   GROUP BY LocationId)   AS ECC_COMMUTERSTU      ON TA.mLoc = ECC_COMMUTERSTU.LocationId 
    --IPASS
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS Ipass_ALL            FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate                                                                                                  GROUP BY LocationId)   AS Ipass_ALL            ON TA.mLoc = Ipass_ALL.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS Ipass_STUDENT        FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND CardType IN (7) AND AreaCode NOT IN (24,25,30,34,38,42,46,50,54,58,62,66,70,74,78,82,86,90,94,98,102,106,110,114) GROUP BY LocationId) AS Ipass_STUDENT ON TA.mLoc = Ipass_STUDENT.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS Ipass_ELDERLY        FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND IdentityType IN (3)      GROUP BY LocationId)   AS Ipass_ELDERLY        ON TA.mLoc = Ipass_ELDERLY.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS Ipass_DISCARD        FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND IdentityType IN (5)      GROUP BY LocationId)   AS Ipass_DISCARD        ON TA.mLoc = Ipass_DISCARD.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS Ipass_DISPARCARD     FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND IdentityType IN (6)      GROUP BY LocationId)   AS Ipass_DISPARCARD     ON TA.mLoc = Ipass_DISPARCARD.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS Ipass_CHICARD        FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND CardType IN (7) AND AreaCode IN (24,25,30,34,38,42,46,50,54,58,62,66,70,74,78,82,86,90,94,98,102,106,110,114) GROUP BY LocationId) AS Ipass_CHICARD ON TA.mLoc = Ipass_CHICARD.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS Ipass_ONEDAY         FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (24) AND PeriodCode IN (1)                                                        GROUP BY LocationId)   AS Ipass_ONEDAY         ON TA.mLoc = Ipass_ONEDAY.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS Ipass_TIMESCARD      FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (24) AND PeriodCode IN (16)                                                       GROUP BY LocationId)   AS Ipass_TIMESCARD      ON TA.mLoc = Ipass_TIMESCARD.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS Ipass_OFFICEMAN      FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (24) AND PeriodCode IN (17)                                                       GROUP BY LocationId)   AS Ipass_OFFICEMAN      ON TA.mLoc = Ipass_OFFICEMAN.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS Ipass_COMMUTERELE_1  FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (28,30)                                                                              GROUP BY LocationId)   AS Ipass_COMMUTERELE    ON TA.mLoc = ECC_COMMUTERELE.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS Ipass_COMMUTERSTU_1  FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (28,30) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND IdentityType in (2) AND CardType in (7)   GROUP BY LocationId)   AS Ipass_COMMUTERSTU    ON TA.mLoc = ECC_COMMUTERSTU.LocationId 
    --ICASH
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS Icash_ALL            FROM #mTempIcash    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate                                                                                                  GROUP BY LocationId)   AS Icash_ALL            ON TA.mLoc = Icash_ALL.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS Icash_STUDENT        FROM #mTempIcash    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND CardType NOT IN (1)      GROUP BY LocationId)   AS Icash_STUDENT        ON TA.mLoc = Icash_STUDENT.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS Icash_ONEDAY         FROM #mTempIcash    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (24) AND PeriodCode IN (1)                                                        GROUP BY LocationId)   AS Icash_ONEDAY         ON TA.mLoc = Icash_ONEDAY.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS Icash_TIMESCARD      FROM #mTempIcash    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (24) AND PeriodCode IN (16)                                                       GROUP BY LocationId)   AS Icash_TIMESCARD      ON TA.mLoc = Icash_TIMESCARD.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS Icash_OFFICEMAN      FROM #mTempIcash    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (24) AND PeriodCode IN (17)                                                       GROUP BY LocationId)   AS Icash_OFFICEMAN      ON TA.mLoc = Icash_OFFICEMAN.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS Icash_COMMUTERELE_1  FROM #mTempIcash    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (28,30)                                                                              GROUP BY LocationId)   AS Icash_COMMUTERELE    ON TA.mLoc = ECC_COMMUTERELE.LocationId 
         LEFT JOIN (SELECT LocationId,     ISNULL(COUNT(*),0) AS Icash_COMMUTERSTU_1  FROM #mTempIcash    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (28,30) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND CardType in (2)          GROUP BY LocationId)   AS Icash_COMMUTERSTU    ON TA.mLoc = ECC_COMMUTERSTU.LocationId 

	ORDER BY LocationId 

DROP TABLE #mTempIcash
DROP TABLE #mTempIpass
DROP TABLE #mTempEcc
DROP TABLE #mTempLocation

";
        public static readonly string Sql_Command_Month_ElectronicTicket_Day = @"--各家票卡運量月報表(天期)
DECLARE @StartStation VARCHAR(10)	--該路線車站起點
DECLARE @EndStation VARCHAR(10)		--該路線車站終點
DECLARE @Openstardate VARCHAR(20)	--搜尋時間開始
DECLARE @Openenddate VARCHAR(20)	--搜尋時間結束
DECLARE @Monthstardate VARCHAR(20)	--開始的月份
DECLARE @Monthenddate VARCHAR(20)	--結束的月份
DECLARE @DateSet VARCHAR(20)		--計算月份用
DECLARE @RunDate VARCHAR(20)		--執行的月份

SET @StartStation = '{0}' 
SET @EndStation = '{1}'
SET @Openstardate = '{2}'
SET @Openenddate = '{3}'
SET @DateSet = CONCAT(DATEPART(YEAR,@Openstardate),'-',DATEPART(MONTH,@Openstardate) , '-01 0' , DATEPART(HOUR,@Openstardate) , ':0' , DATEPART(MINUTE,@Openstardate) , ':0' , DATEPART(SECOND,@Openstardate))
SET @Monthenddate = DATEADD(MONTH,1,@DateSet)
SET @Monthstardate = DATEADD(MONTH,-1,@Monthenddate)
SET @RunDate = @Monthstardate

SELECT LocationId as mLoc
INTO #mTempLocation
FROM Parm081_LocList
WHERE LocationId >= @StartStation AND LocationId <= @EndStation AND CurrentType IN (1)

SELECT *
INTO #mTempEcc
FROM Txn_Exit JOIN #mTempLocation on LocationId = mLoc
WHERE TxnDT >= @Monthstardate AND TxnDT < @Monthenddate AND IssueCode IN (2) 

SELECT *
INTO #mTempIpass
FROM Txn_Exit JOIN #mTempLocation on LocationId = mLoc
WHERE TxnDT >= @Monthstardate AND TxnDT < @Monthenddate AND IssueCode IN (9) 

SELECT *
INTO #mTempIcash
FROM Txn_Exit JOIN #mTempLocation on LocationId = mLoc
WHERE TxnDT >= @Monthstardate AND TxnDT < @Monthenddate AND IssueCode IN (11) 

CREATE TABLE #mTemplate (
    mYear INT,
    mMonth INT,
    mDay INT,
    Ecc_ElectronicCard          INT,    --悠遊卡_普通卡
    Ecc_StudentCard             INT,    --悠遊卡_學生卡
    Ecc_ElderlyCard             INT,    --悠遊卡_敬老卡
    Ecc_DisabledCard            INT,    --悠遊卡_愛心卡
    Ecc_DisabledPartherCard     INT,    --悠遊卡_愛陪卡
    Ecc_ChildCard               INT,    --悠遊卡_兒童卡
    Ecc_Commuter_ElectronicAll  INT,    --悠遊卡_定期票_補通卡
    Ecc_Commuter_StudentCard    INT,    --悠遊卡_定期票_學生卡
    Ipass_ElectronicCard        INT,    --一卡通_普通卡
    Ipass_StudentCard           INT,    --一卡通_學生卡
    Ipass_ElderlyCard           INT,    --一卡通_敬老卡
    Ipass_DisabledCard          INT,    --一卡通_愛心卡
    Ipass_DisabledPartherCard   INT,    --一卡通_愛陪卡
    Ipass_ChildCard             INT,    --一卡通_兒童卡
    Icash_ElectronicCard        INT,    --愛金卡_普通卡
    Icash_WelfareCard           INT,    --愛金卡_優待卡
    Credit_M_Exit               INT,    --信用卡_萬事達_出站
    Credit_I_Exit               INT,    --信用卡_銀聯_出站
    Credit_D_Exit               INT,    --信用卡_發現卡_出站
    QrCode_P_Exit               INT,    --乘車碼_一卡通_出站
    QrCode_T_Exit               INT,    --乘車碼_台新_出站

)

WHILE(@RunDate < @Monthenddate)
BEGIN
    INSERT INTO #mTemplate     
    SELECT 
        DATEPART(YEAR,@RunDate) AS mYear, DATEPART(MONTH,@RunDate) AS mMonth, DATEPART(DAY,@RunDate) AS mDay, 
    --ECC
        (ISNULL(ECC_ALL,0) - ISNULL(ECC_STUDENT,0) - ISNULL(ECC_ELDERLY,0) - ISNULL(ECC_DISCARD,0) - ISNULL(ECC_DISPARCARD,0) - ISNULL(ECC_CHICARD,0) - ISNULL(ECC_ONEDAY,0) - ISNULL(ECC_OFFICEMAN,0) - ISNULL(ECC_COMMUTERELE_1,0)) AS Ecc_ElectronicCard, 
        ISNULL(ECC_STUDENT,0)         AS Ecc_StudentCard, 
        ISNULL(ECC_ELDERLY,0)         AS Ecc_ElderlyCard, 
        ISNULL(ECC_DISCARD,0)         AS Ecc_DisabledCard, 
        ISNULL(ECC_DISPARCARD,0)      AS Ecc_DisabledPartherCard, 
        ISNULL(ECC_CHICARD,0)         AS Ecc_ChildCard, 
        (ISNULL(ECC_COMMUTERELE_1,0) - ISNULL(ECC_COMMUTERSTU_1,0)) AS Ecc_Commuter_ElectronicAll_1, 
        ISNULL(ECC_COMMUTERSTU_1,0)     AS Ecc_Commuter_StudentCard_1, 
    --IPASS
        --Ipass_ALL as Ipass_ALL, 
        (ISNULL(Ipass_ALL,0) - ISNULL(Ipass_STUDENT,0) - ISNULL(Ipass_ELDERLY,0) - ISNULL(Ipass_DISCARD,0) - ISNULL(Ipass_DISPARCARD,0) - ISNULL(Ipass_CHICARD,0) - ISNULL(Ipass_ONEDAY,0) - ISNULL(Ipass_OFFICEMAN,0)) AS Ipass_ElectronicCard, 
        ISNULL(Ipass_STUDENT,0)       AS Ipass_StudentCard, 
        ISNULL(Ipass_ELDERLY,0)       AS Ipass_ElderlyCard, 
        ISNULL(Ipass_DISCARD,0)       AS Ipass_DisabledCard, 
        ISNULL(Ipass_DISPARCARD,0)    AS Ipass_DisabledPartherCard, 
        ISNULL(Ipass_CHICARD,0)       AS Ipass_ChildCard, 
    --ICASH
        (ISNULL(Icash_ALL,0) - ISNULL(Icash_STUDENT,0) - ISNULL(Icash_ONEDAY,0) - ISNULL(Icash_OFFICEMAN,0)) AS Icash_ElectronicCard, 
        ISNULL(Icash_STUDENT,0)       AS Icash_WelfareCard, 
    --Cridet
        0 AS Credit_M_Exit,
        0 AS Credit_I_Exit,
        0 AS Credit_D_Exit,
        0 AS QrCode_P_Exit,
        0 AS QrCode_T_Exit
    FROM 
    --ECC
                   (SELECT ISNULL(COUNT(*),0) AS ECC_ALL              FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate)                                                                                                                                                                                           ) AS ECC_ALL              
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS ECC_STUDENT          FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND PersonalProfile IN (5)                                                                                            ) AS ECC_STUDENT       ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS ECC_ELDERLY          FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND PersonalProfile IN (1,2)                                                                                          ) AS ECC_ELDERLY       ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS ECC_DISCARD          FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND PersonalProfile IN (3)                                                                                            ) AS ECC_DISCARD       ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS ECC_DISPARCARD       FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND PersonalProfile IN (4)                                                                                            ) AS ECC_DISPARCARD    ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS ECC_CHICARD          FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND PersonalProfile IN (8) AND AreaCode IN (1,2)                                                                      ) AS ECC_CHICARD       ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS ECC_ONEDAY           FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (24) AND PeriodCode IN (1)                                                                                                                                                 ) AS ECC_ONEDAY        ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS ECC_TIMESCARD        FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (24) AND PeriodCode IN (16)                                                                                                                                                ) AS ECC_TIMESCARD     ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS ECC_OFFICEMAN        FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (24) AND PeriodCode IN (17)                                                                                                                                                ) AS ECC_OFFICEMAN     ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS ECC_COMMUTERELE_1    FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (28,30)                                                                                                                                                                    ) AS ECC_COMMUTERELE   ON 1=1 
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS ECC_COMMUTERSTU_1    FROM #mTempEcc      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (28,30) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND PersonalProfile IN (5)                                                                                         ) AS ECC_COMMUTERSTU   ON 1=1 
    --IPASS                                                                                                                                                                                                                                                                                                                                                                                                                   
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS Ipass_ALL            FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate)                                                                                                                                                                                           ) AS Ipass_ALL         ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS Ipass_STUDENT        FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND CardType IN (7) AND AreaCode NOT IN (24,25,30,34,38,42,46,50,54,58,62,66,70,74,78,82,86,90,94,98,102,106,110,114) ) AS Ipass_STUDENT     ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS Ipass_ELDERLY        FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND IdentityType IN (3)                                                                                               ) AS Ipass_ELDERLY     ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS Ipass_DISCARD        FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND IdentityType IN (5)                                                                                               ) AS Ipass_DISCARD     ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS Ipass_DISPARCARD     FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND IdentityType IN (6)                                                                                               ) AS Ipass_DISPARCARD  ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS Ipass_CHICARD        FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND CardType IN (7) AND AreaCode IN (24,25,30,34,38,42,46,50,54,58,62,66,70,74,78,82,86,90,94,98,102,106,110,114)     ) AS Ipass_CHICARD     ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS Ipass_ONEDAY         FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (24) AND PeriodCode IN (1)                                                                                                                                                 ) AS Ipass_ONEDAY      ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS Ipass_TIMESCARD      FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (24) AND PeriodCode IN (16)                                                                                                                                                ) AS Ipass_TIMESCARD   ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS Ipass_OFFICEMAN      FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (24) AND PeriodCode IN (17)                                                                                                                                                ) AS Ipass_OFFICEMAN   ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS Ipass_COMMUTERELE_1  FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (28,30)                                                                                                                                                                    ) AS Ipass_COMMUTERELE ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS Ipass_COMMUTERSTU_1  FROM #mTempIpass    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (28,30) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND IdentityType in (2) AND CardType in (7)                                                                        ) AS Ipass_COMMUTERSTU ON 1=1
    --ICASH                                                                                                                                                                                                                                                                                                                                                                                              
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS Icash_ALL            FROM #mTempIcash    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate)                                                                                                                                                                                           ) AS Icash_ALL         ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS Icash_STUDENT        FROM #mTempIcash    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (22) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND CardType NOT IN (1)                                                                                               ) AS Icash_STUDENT     ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS Icash_ONEDAY         FROM #mTempIcash    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (24) AND PeriodCode IN (1)                                                                                                                                                 ) AS Icash_ONEDAY      ON 1=1
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS Icash_TIMESCARD      FROM #mTempIcash    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (24) AND PeriodCode IN (16)                                                                                                                                                ) AS Icash_TIMESCARD   ON 1=1 
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS Icash_OFFICEMAN      FROM #mTempIcash    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (24) AND PeriodCode IN (17)                                                                                                                                                ) AS Icash_OFFICEMAN   ON 1=1 
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS Icash_COMMUTERELE_1  FROM #mTempIcash    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (28,30)                                                                                                                                                                    ) AS Icash_COMMUTERELE ON 1=1 
         LEFT JOIN (SELECT ISNULL(COUNT(*),0) AS Icash_COMMUTERSTU_1  FROM #mTempIcash    WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(day,1,@RunDate) AND TxnType IN (28,30) AND DATEDIFF(MINUTE,TxnDT,IdentityExpiryDT) > 0 AND CardType in (2)                                                                                                ) AS Icash_COMMUTERSTU ON 1=1 
    SET @RunDate = DATEADD(day,1,@RunDate)
END

SELECT *
FROM #mTemplate
ORDER BY mYear,mMonth,mDay

DROP TABLE #mTempIcash
DROP TABLE #mTempIpass
DROP TABLE #mTempEcc
DROP TABLE #mTempLocation
DROP TABLE #mTemplate

";
        public static readonly string Sql_Command_Month_OwnTicketVolume_Station = @"--自有票卡運量月報表(車站)
DECLARE @StartStation VARCHAR(10)	--該路線車站起點
DECLARE @EndStation VARCHAR(10)		--該路線車站終點
DECLARE @Openstardate VARCHAR(20)	--搜尋時間開始
DECLARE @Openenddate VARCHAR(20)	--搜尋時間結束
DECLARE @Monthstardate VARCHAR(20)	--開始的月份
DECLARE @Monthenddate VARCHAR(20)	--結束的月份
DECLARE @DateSet VARCHAR(20)		--計算月份用
DECLARE @RunDate VARCHAR(20)		--執行的月份

SET @StartStation = '{0}' 
SET @EndStation = '{1}'
SET @Openstardate = '{2}'
SET @Openenddate = '{3}'
SET @DateSet = CONCAT(DATEPART(YEAR,@Openstardate),'-',DATEPART(MONTH,@Openstardate) , '-01 0' , DATEPART(HOUR,@Openstardate) , ':0' , DATEPART(MINUTE,@Openstardate) , ':0' , DATEPART(SECOND,@Openstardate))
SET @Monthenddate = DATEADD(MONTH,1,@DateSet)
SET @Monthstardate = DATEADD(MONTH,-1,@Monthenddate)
SET @RunDate = @Monthstardate

SELECT LocationId as mLoc
INTO #mTempLocation
FROM Parm081_LocList
WHERE LocationId >= @StartStation AND LocationId <= @EndStation AND CurrentType IN (1)

SELECT *
INTO #mTempTicket
FROM Txn_SellSpecialTicket JOIN #mTempLocation on LocationId = mLoc
WHERE TxnDT >= @Monthstardate AND TxnDT < @Monthenddate

SELECT *
INTO #mTempReTicket
FROM Txn_RefundSpeTkt JOIN #mTempLocation on LocationId = mLoc
WHERE TxnDT >= @Monthstardate AND TxnDT < @Monthenddate

SELECT *
INTO #mTempOffice
FROM Txn_Exit JOIN #mTempLocation on LocationId = mLoc
WHERE TxnDT >= @Monthstardate AND TxnDT < @Monthenddate AND IssueCode IN (9) AND TxnType IN (24) AND PeriodCode = 17

SELECT *
INTO #mTempOneDay
FROM Txn_Exit JOIN #mTempLocation on LocationId = mLoc
WHERE TxnDT >= @Monthstardate AND TxnDT < @Monthenddate AND IssueCode IN (9) AND TxnType IN (24) AND PeriodCode = 1

    
SELECT 
    mLoc as LocationId 
    ,ISNULL(OneTimeGeneral   , 0) - ISNULL(OneTimeReGeneral, 0)     AS OneTimeGeneral 
    ,ISNULL(OneTimeConcession, 0) - ISNULL(OneTimeReConcession, 0)  AS OneTimeConcession 
    ,ISNULL(OneTimeBike      , 0) - ISNULL(OneTimeReBike, 0)        AS OneTimeBike 
    ,ISNULL(OfficeCard, 0)                               AS OfficeCard 
    ,ISNULL(OneDayTicket, 0)                             AS OneDayTicket
    --,0 AS MenGerneral,
    --,0 AS MenConcession,
    --,0 AS MenBike,
    --,0 AS MenOneDayCount,
    --,0 AS MenOneDayTimes,
    --,0 AS MenOneDayCount2,
    --,0 AS MenOneDayTimes2,
    --,0 AS GroupCount,
    --,0 AS GroupTimes
FROM 
              (SELECT mLoc FROM #mTempLocation) AS TA
    LEFT JOIN (SELECT LocationId,ISNULL(COUNT(*),0) AS OneTimeGeneral        FROM #mTempTicket   WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND CardType IN (1) GROUP BY LocationId)    AS OneTimeGeneral       ON TA.mLoc = OneTimeGeneral.LocationId 
    LEFT JOIN (SELECT LocationId,ISNULL(COUNT(*),0) AS OneTimeConcession     FROM #mTempTicket   WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND CardType IN (3) GROUP BY LocationId)    AS OneTimeConcession    ON TA.mLoc = OneTimeConcession.LocationId 
    LEFT JOIN (SELECT LocationId,ISNULL(COUNT(*),0) AS OneTimeBike           FROM #mTempTicket   WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND CardType IN (2) GROUP BY LocationId)    AS OneTimeBike          ON TA.mLoc = OneTimeBike.LocationId 
    LEFT JOIN (SELECT LocationId,ISNULL(COUNT(*),0) AS OneTimeReGeneral      FROM #mTempReTicket WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND CardType IN (1) GROUP BY LocationId)    AS OneTimeReGeneral     ON TA.mLoc = OneTimeReGeneral.LocationId 
    LEFT JOIN (SELECT LocationId,ISNULL(COUNT(*),0) AS OneTimeReConcession   FROM #mTempReTicket WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND CardType IN (3) GROUP BY LocationId)    AS OneTimeReConcession  ON TA.mLoc = OneTimeReConcession.LocationId 
    LEFT JOIN (SELECT LocationId,ISNULL(COUNT(*),0) AS OneTimeReBike         FROM #mTempReTicket WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND CardType IN (2) GROUP BY LocationId)    AS OneTimeReBike        ON TA.mLoc = OneTimeReBike.LocationId 
    LEFT JOIN (SELECT LocationId,ISNULL(COUNT(*),0) AS OfficeCard            FROM #mTempOffice   WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate                     GROUP BY LocationId)    AS OfficeCard           ON TA.mLoc = OfficeCard.LocationId 
    LEFT JOIN (SELECT LocationId,ISNULL(COUNT(*),0) AS OneDayTicket          FROM #mTempOneDay   WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate                     GROUP BY LocationId)    AS OneDayTicket         ON TA.mLoc = OneDayTicket.LocationId 
ORDER BY LocationId

select *
from 
	(select COUNT(*) as Ticket from #mTempTicket ) as t1 
	left join (select COUNT(*) as rTicket from #mTempReTicket) as T2 on 1=1

DROP TABLE #mTempOneDay
DROP TABLE #mTempOffice
DROP TABLE #mTempReTicket
DROP TABLE #mTempTicket
DROP TABLE #mTempLocation

";
        public static readonly string Sql_Command_Month_OwnTicketVolume_Day = @"--自有票卡運量月報表(天期)
DECLARE @StartStation VARCHAR(10)	--該路線車站起點
DECLARE @EndStation VARCHAR(10)		--該路線車站終點
DECLARE @Openstardate VARCHAR(20)	--搜尋時間開始
DECLARE @Openenddate VARCHAR(20)	--搜尋時間結束
DECLARE @Monthstardate VARCHAR(20)	--開始的月份
DECLARE @Monthenddate VARCHAR(20)	--結束的月份
DECLARE @DateSet VARCHAR(20)		--計算月份用
DECLARE @RunDate VARCHAR(20)		--執行的月份

SET @StartStation = '{0}' 
SET @EndStation = '{1}'
SET @Openstardate = '{2}'
SET @Openenddate = '{3}'
SET @DateSet = CONCAT(DATEPART(YEAR,@Openstardate),'-',DATEPART(MONTH,@Openstardate) , '-01 0' , DATEPART(HOUR,@Openstardate) , ':0' , DATEPART(MINUTE,@Openstardate) , ':0' , DATEPART(SECOND,@Openstardate))
SET @Monthenddate = DATEADD(MONTH,1,@DateSet)
SET @Monthstardate = DATEADD(MONTH,-1,@Monthenddate)
SET @RunDate = @Monthstardate

SELECT LocationId as mLoc
INTO #mTempLocation
FROM Parm081_LocList
WHERE LocationId >= @StartStation AND LocationId <= @EndStation AND CurrentType IN (1)

SELECT *
INTO #mTempTicket
FROM Txn_SellSpecialTicket JOIN #mTempLocation on LocationId = mLoc
WHERE TxnDT >= @Monthstardate AND TxnDT < @Monthenddate

SELECT *
INTO #mTempReTicket
FROM Txn_RefundSpeTkt JOIN #mTempLocation on LocationId = mLoc
WHERE TxnDT >= @Monthstardate AND TxnDT < @Monthenddate

SELECT *
INTO #mTempOffice
FROM Txn_Exit JOIN #mTempLocation on LocationId = mLoc
WHERE TxnDT >= @Monthstardate AND TxnDT < @Monthenddate AND IssueCode IN (9) AND TxnType IN (24) AND PeriodCode = 17

SELECT *
INTO #mTempOneDay
FROM Txn_Exit JOIN #mTempLocation on LocationId = mLoc
WHERE TxnDT >= @Monthstardate AND TxnDT < @Monthenddate AND IssueCode IN (9) AND TxnType IN (24) AND PeriodCode = 1

CREATE TABLE #mTemplate (
    mYear INT
    ,mMonth INT
    ,mDay INT
    ,OneTimeGeneral     INT    --系統單程票_一般
    ,OneTimeConcession  INT    --系統單程票_優待
    ,OneTimeBike        INT    --系統單程票_自行車
    ,OfficeCard         INT    --員工卡
    ,OneDayTicket       INT    --系統一日票
    --,MenGerneral        INT    --人工單程票_一般
    --,MenConcession      INT    --人工單程票_優待
    --,MenBike            INT    --人工單程票_自行車
    --,MenOneDayCount     INT    --人工單程票_一日票(張)
    --,MenOneDayTimes     INT    --人工單程票_一日票(人次)
    --,MenOneDayCount2    INT    --人工單程票_一日票2(張)
    --,MenOneDayTimes2    INT    --人工單程票_一日票2(人次)
    --,GroupCount         INT    --人工單程票_團體票(張)
    --,GroupTimes         INT    --人工單程票_團體票(人次)
)

WHILE(@RunDate < @Monthenddate)
BEGIN
    INSERT INTO #mTemplate     
    SELECT DATEPART(YEAR,@RunDate) AS mYear, DATEPART(MONTH,@RunDate) AS mMonth, DATEPART(DAY,@RunDate) AS mDay
        ,ISNULL(OneTimeGeneral    - OneTimeReGeneral, 0)     AS OneTimeGeneral 
        ,ISNULL(OneTimeConcession - OneTimeReConcession, 0)  AS OneTimeConcession 
        ,ISNULL(OneTimeBike       - OneTimeReBike, 0)        AS OneTimeBike 
        ,ISNULL(OfficeCard, 0)                               AS OfficeCard 
        ,ISNULL(OneDayTicket, 0)                             AS OneDayTicket
      --,0 AS MenGerneral,
      --,0 AS MenConcession,
      --,0 AS MenBike,
      --,0 AS MenOneDayCount,
      --,0 AS MenOneDayTimes,
      --,0 AS MenOneDayCount2,
      --,0 AS MenOneDayTimes2,
      --,0 AS GroupCount,
      --,0 AS GroupTimes
    From
        (SELECT ISNULL(COUNT(*),0) AS OneTimeGeneral        FROM #mTempTicket   WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND CardType IN (1))    AS OneTimeGeneral               LEFT JOIN
        (SELECT ISNULL(COUNT(*),0) AS OneTimeConcession     FROM #mTempTicket   WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND CardType IN (3))    AS OneTimeConcession    ON 1=1  LEFT JOIN
        (SELECT ISNULL(COUNT(*),0) AS OneTimeBike           FROM #mTempTicket   WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND CardType IN (2))    AS OneTimeBike          ON 1=1  LEFT JOIN
        (SELECT ISNULL(COUNT(*),0) AS OneTimeReGeneral      FROM #mTempReTicket WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND CardType IN (1))    AS OneTimeReGeneral     ON 1=1  LEFT JOIN
        (SELECT ISNULL(COUNT(*),0) AS OneTimeReConcession   FROM #mTempReTicket WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND CardType IN (3))    AS OneTimeReConcession  ON 1=1  LEFT JOIN
        (SELECT ISNULL(COUNT(*),0) AS OneTimeReBike         FROM #mTempReTicket WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND CardType IN (2))    AS OneTimeReBike        ON 1=1  LEFT JOIN
        (SELECT ISNULL(COUNT(*),0) AS OfficeCard            FROM #mTempOffice   WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate))                        AS OfficeCard           ON 1=1  LEFT JOIN
        (SELECT ISNULL(COUNT(*),0) AS OneDayTicket          FROM #mTempOneDay   WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate))                        AS OneDayTicket         ON 1=1  

    SET @RunDate = DATEADD(day,1,@RunDate)
END

select *
from #mTemplate
order by mYear,mMonth,mDay

DROP TABLE #mTempOneDay
DROP TABLE #mTempOffice
DROP TABLE #mTempReTicket
DROP TABLE #mTempTicket
DROP TABLE #mTempLocation
DROP TABLE #mTemplate

";
        public static readonly string Sql_Command_Month_TrafficAmount_Station = @"--營收月報表(車站)
DECLARE @StartStation VARCHAR(10)	--該路線車站起點
DECLARE @EndStation VARCHAR(10)		--該路線車站終點
DECLARE @Openstardate VARCHAR(20)	--搜尋時間開始
DECLARE @Openenddate VARCHAR(20)	--搜尋時間結束
DECLARE @Monthstardate VARCHAR(20)	--開始的月份
DECLARE @Monthenddate VARCHAR(20)	--結束的月份
DECLARE @DateSet VARCHAR(20)		--計算月份用
DECLARE @RunDate VARCHAR(20)		--執行的月份

SET @StartStation = '{0}' 
SET @EndStation = '{1}'
SET @Openstardate = '{2}'
SET @Openenddate = '{3}'
SET @DateSet = CONCAT(DATEPART(YEAR,@Openstardate),'-',DATEPART(MONTH,@Openstardate) , '-01 0' , DATEPART(HOUR,@Openstardate) , ':0' , DATEPART(MINUTE,@Openstardate) , ':0' , DATEPART(SECOND,@Openstardate))
SET @Monthenddate = DATEADD(MONTH,1,@DateSet)
SET @Monthstardate = DATEADD(MONTH,-1,@Monthenddate)
SET @RunDate = @Monthstardate

SELECT 
    TA.LocationId,
    (ISNULL(T1.ECCTxnAmt,0) + ISNULL(T8.TxnAmtCsc,0) + ISNULL(T9.TxnAmtDed,0)) AS ECCTxnAmt,
    ISNULL(T2.IPASSTxnAmt,0) AS IPASSTxnAmt,
    ISNULL(T3.ICASHTxnAmt,0) AS ICASHTxnAmt,
    0 AS CreditCard,
    0 AS QrCode,
    ISNULL(T4.SOneTkt_Amt,0) - ISNULL(T10.TxnAmtRe,0)AS SOneTkt_Amt,
    ISNULL(T5.SOneTkt_Amt_discount,0) AS SOneTkt_Amt_discount,
    ISNULL(T7.SOneTkt_bike_Amt,0) AS SOneTkt_bike_Amt,
    0 AS OneDayTicket1,
    0 AS OneDayTicket2,
    0 AS GroupTicket,
    0 AS OpenTrafficeTicket,
    0 AS ReTicket,
    0 AS AfcTackOut,
    0 AS VavmReFound,
    0 AS SaleCardReFound,
    0 AS PvReFound
FROM ( 
               SELECT LocationId  FROM Parm081_LocList WHERE CurrentType IN (1) AND LocationId BETWEEN @StartStation AND @EndStation) AS TA 
    LEFT JOIN (SELECT LocationId, ISNULL(SUM(TxnAmt),0) AS ECCTxnAmt               FROM Txn_Exit               WHERE TxnDT >= @Openstardate and TxnDT < @Openenddate and TxnType in(22) and IssueCode = 2  GROUP BY LocationId ) AS T1 ON TA.LocationId = T1.LocationId 
    LEFT JOIN (SELECT LocationId, ISNULL(SUM(TxnAmt),0) AS IPASSTxnAmt             FROM Txn_Exit               WHERE TxnDT >= @Openstardate and TxnDT < @Openenddate and TxnType in(22) and IssueCode = 9  GROUP BY LocationId ) AS T2 ON TA.LocationId = T2.LocationId 
    LEFT JOIN (SELECT LocationId, ISNULL(SUM(TxnAmt),0) AS ICASHTxnAmt             FROM Txn_Exit               WHERE TxnDT >= @Openstardate and TxnDT < @Openenddate and TxnType in(22) and IssueCode = 11 GROUP BY LocationId ) AS T3 ON TA.LocationId = T3.LocationId 
    LEFT JOIN (SELECT LocationId, ISNULL(SUM(TxnAmt),0) AS SOneTkt_Amt             FROM Txn_SellSpecialTicket  WHERE TxnDT >= @Openstardate and TxnDT < @Openenddate and CardType in(1)                    GROUP BY LocationId ) AS T4 ON TA.LocationId = T4.LocationId 
    LEFT JOIN (SELECT LocationId, ISNULL(SUM(TxnAmt),0) AS SOneTkt_Amt_discount    FROM Txn_SellSpecialTicket  WHERE TxnDT >= @Openstardate and TxnDT < @Openenddate and CardType in(3)                    GROUP BY LocationId ) AS T5 ON TA.LocationId = T5.LocationId 
    LEFT JOIN (SELECT LocationId, ISNULL(SUM(TxnAmt),0) AS SOneTkt_bike_Amt        FROM Txn_SellSpecialTicket  WHERE TxnDT >= @Openstardate and TxnDT < @Openenddate and CardType in(2)                    GROUP BY LocationId ) AS T7 ON TA.LocationId = T7.LocationId 
    LEFT JOIN (SELECT LocationId, ISNULL(SUM(TxnAmt),0) AS TxnAmtCsc               FROM Txn_ExcessByCSC        WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate                                       GROUP BY LocationId ) AS T8 ON TA.LocationId = T8.LocationId 
    LEFT JOIN (SELECT LocationId, ISNULL(SUM(TxnAmt),0) AS TxnAmtDed               FROM Txn_ExceptDeduct       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate                                       GROUP BY LocationId ) AS T9 ON TA.LocationId = T9.LocationId 
    LEFT JOIN (SELECT LocationId, ISNULL(SUM(TxnAmt),0) AS TxnAmtRe                FROM Txn_RefundSpeTkt       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType IN (97,98,99)             GROUP BY LocationId ) AS T10 ON TA.LocationId = T10.LocationId 
ORDER BY LocationId
";
        public static readonly string Sql_Command_Month_TrafficAmount_Day = @"--營收月報表(天期)
DECLARE @StartStation VARCHAR(10)	--該路線車站起點
DECLARE @EndStation VARCHAR(10)		--該路線車站終點
DECLARE @Openstardate VARCHAR(20)	--搜尋時間開始
DECLARE @Openenddate VARCHAR(20)	--搜尋時間結束
DECLARE @Monthstardate VARCHAR(20)	--開始的月份
DECLARE @Monthenddate VARCHAR(20)	--結束的月份
DECLARE @DateSet VARCHAR(20)		--計算月份用
DECLARE @RunDate VARCHAR(20)		--執行的月份

SET @StartStation = '{0}' 
SET @EndStation = '{1}'
SET @Openstardate = '{2}'
SET @Openenddate = '{3}'
SET @DateSet = CONCAT(DATEPART(YEAR,@Openstardate),'-',DATEPART(MONTH,@Openstardate) , '-01 0' , DATEPART(HOUR,@Openstardate) , ':0' , DATEPART(MINUTE,@Openstardate) , ':0' , DATEPART(SECOND,@Openstardate))
SET @Monthenddate = DATEADD(MONTH,1,@DateSet)
SET @Monthstardate = DATEADD(MONTH,-1,@Monthenddate)
SET @RunDate = @Monthstardate

SELECT LocationId as mLoc
INTO #mTempLocation
FROM Parm081_LocList
WHERE LocationId >= @StartStation AND LocationId <= @EndStation AND CurrentType IN (1)

CREATE TABLE #mTemplate (
    mYear                INT,
    mMonth               INT,
    mDay                 INT,
    ECCAmt               INT,
    IpassAmt             INT,
    IcashAmt             INT,
    CreditCard           INT,
    QrCode               INT,
    SOneTkt_Amt          INT,
    SOneTkt_Amt_discount INT,
    SOneTkt_bike_Amt     INT,
    OneDayTicket1		 INT,
    OneDayTicket2		 INT,
    GroupTicket			 INT,
    OpenTrafficeTicket	 INT,
    ReTicket			 INT,
    AfcTackOut			 INT,
    VavmReFound			 INT,
    SaleCardReFound		 INT,
    PvReFound			 INT
)

WHILE(@RunDate < @Monthenddate)
BEGIN
    INSERT INTO #mTemplate     
    SELECT 
        DATEPART(YEAR,@RunDate) AS mYear, DATEPART(MONTH,@RunDate) AS mMonth, DATEPART(DAY,@RunDate) AS mDay, 
        ISNULL(SUM(ECCTxnAmt),0) + ISNULL(SUM(ECCTxnAmtCsc),0) + ISNULL(SUM(ECCTxnAmtDed),0) AS ECCTxnAmt, 
        ISNULL(SUM(IPASSTxnAmt),0) + ISNULL(SUM(IPASSTxnAmtCsc),0) + ISNULL(SUM(IPASSTxnAmtDed),0) AS IpassAmt, 
        ISNULL(SUM(ICASHTxnAmt),0) + ISNULL(SUM(ICASHTxnAmtCsc),0) + ISNULL(SUM(ICASHTxnAmtDed),0) AS IcashAmt, 
        0 AS CreditCard,
        0 AS QrCode,
        ISNULL(SUM(SOneTkt_Amt),0) - ISNULL(SUM(TxnAmtRe1),0) AS SOneTkt_Amt, 
        ISNULL(SUM(SOneTkt_Amt_discount),0) - ISNULL(SUM(TxnAmtRe2),0) AS SOneTkt_Amt_discount, 
        ISNULL(SUM(SOneTkt_bike_Amt),0) - ISNULL(SUM(TxnAmtRe3),0) AS SOneTkt_bike_Amt ,
    0 AS OneDayTicket1,
    0 AS OneDayTicket2,
    0 AS GroupTicket,
    0 AS OpenTrafficeTicket,
    0 AS ReTicket,
    0 AS AfcTackOut,
    0 AS VavmReFound,
    0 AS SaleCardReFound,
    0 AS PvReFound
FROM
              (SELECT ISNULL(SUM(TxnAmt),0) AS ECCTxnAmt            FROM Txn_Exit               JOIN #mTempLocation ON LocationId   = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(DAY,1,@RunDate) AND TxnType  IN (22) AND IssueCode = 2  GROUP BY IssueCode) AS T10
    LEFT JOIN (SELECT ISNULL(SUM(TxnAmt),0) AS ECCTxnAmtCsc         FROM Txn_ExcessByCSC        JOIN #mTempLocation ON LocationId   = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(DAY,1,@RunDate) AND IssueCode = 2                       GROUP BY IssueCode) AS T11 ON 1 = 1 
    LEFT JOIN (SELECT ISNULL(SUM(TxnAmt),0) AS ECCTxnAmtDed         FROM Txn_ExceptDeduct       JOIN #mTempLocation ON LocationId   = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(DAY,1,@RunDate) AND IssueCode = 2                       GROUP BY IssueCode) AS T12 ON 1 = 1  
    LEFT JOIN (SELECT ISNULL(SUM(TxnAmt),0) AS IPASSTxnAmt          FROM Txn_Exit               JOIN #mTempLocation ON LocationId   = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(DAY,1,@RunDate) AND TxnType  IN (22) AND IssueCode = 9  GROUP BY IssueCode) AS T20 ON 1 = 1  
    LEFT JOIN (SELECT ISNULL(SUM(TxnAmt),0) AS IPASSTxnAmtCsc       FROM Txn_ExcessByCSC        JOIN #mTempLocation ON LocationId   = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(DAY,1,@RunDate) AND IssueCode = 9                       GROUP BY IssueCode) AS T21 ON 1 = 1   
    LEFT JOIN (SELECT ISNULL(SUM(TxnAmt),0) AS IPASSTxnAmtDed       FROM Txn_ExceptDeduct       JOIN #mTempLocation ON LocationId   = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(DAY,1,@RunDate) AND IssueCode = 9                       GROUP BY IssueCode) AS T22 ON 1 = 1  
    LEFT JOIN (SELECT ISNULL(SUM(TxnAmt),0) AS ICASHTxnAmt          FROM Txn_Exit               JOIN #mTempLocation ON LocationId   = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(DAY,1,@RunDate) AND TxnType  IN (22) AND IssueCode = 11 GROUP BY IssueCode) AS T30 ON 1 = 1  
    LEFT JOIN (SELECT ISNULL(SUM(TxnAmt),0) AS ICASHTxnAmtCsc       FROM Txn_ExcessByCSC        JOIN #mTempLocation ON LocationId   = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(DAY,1,@RunDate) AND IssueCode = 11                      GROUP BY IssueCode) AS T31 ON 1 = 1   
    LEFT JOIN (SELECT ISNULL(SUM(TxnAmt),0) AS ICASHTxnAmtDed       FROM Txn_ExceptDeduct       JOIN #mTempLocation ON LocationId   = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(DAY,1,@RunDate) AND IssueCode = 11                      GROUP BY IssueCode) AS T32 ON 1 = 1  
    LEFT JOIN (SELECT ISNULL(SUM(TxnAmt),0) AS SOneTkt_Amt          FROM Txn_SellSpecialTicket  JOIN #mTempLocation ON LocationId   = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(DAY,1,@RunDate) AND CardType IN (1)                     GROUP BY CardType ) AS T40 ON 1 = 1  
    LEFT JOIN (SELECT ISNULL(SUM(TxnAmt),0) AS TxnAmtRe1            FROM Txn_RefundSpeTkt       JOIN #mTempLocation ON LocationId   = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(DAY,1,@RunDate) AND CardType IN (1)                     GROUP BY CardType ) AS T41 ON 1 = 1  
    LEFT JOIN (SELECT ISNULL(SUM(TxnAmt),0) AS SOneTkt_Amt_discount FROM Txn_SellSpecialTicket  JOIN #mTempLocation ON LocationId   = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(DAY,1,@RunDate) AND CardType IN (3)                     GROUP BY CardType ) AS T50 ON 1 = 1   
    LEFT JOIN (SELECT ISNULL(SUM(TxnAmt),0) AS TxnAmtRe2            FROM Txn_RefundSpeTkt       JOIN #mTempLocation ON LocationId   = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(DAY,1,@RunDate) AND CardType IN (3)                     GROUP BY CardType ) AS T51 ON 1 = 1  
    LEFT JOIN (SELECT ISNULL(SUM(TxnAmt),0) AS SOneTkt_bike_Amt     FROM Txn_SellSpecialTicket  JOIN #mTempLocation ON LocationId   = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(DAY,1,@RunDate) AND CardType IN (2)                     GROUP BY CardType ) AS T60 ON 1 = 1  
    LEFT JOIN (SELECT ISNULL(SUM(TxnAmt),0) AS TxnAmtRe3            FROM Txn_RefundSpeTkt       JOIN #mTempLocation ON LocationId   = mLoc  WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDT < DATEADD(DAY,1,@RunDate) AND CardType IN (2)                     GROUP BY CardType ) AS T61 ON 1 = 1  


    SET @RunDate = DATEADD(day,1,@RunDate)
END

SELECT * FROM #mTemplate ORDER BY mYear,mMonth,mDay

DROP TABLE #mTempLocation
DROP TABLE #mTemplate
";
        public static readonly string Sql_Command_Month_EquipAmount_Station = @"--設備營收日報表
DECLARE @StartStation VARCHAR(10)	--該路線車站起點
DECLARE @EndStation VARCHAR(10)		--該路線車站終點
DECLARE @Openstardate VARCHAR(20)	--搜尋時間開始
DECLARE @Openenddate VARCHAR(20)	--搜尋時間結束
DECLARE @Monthstardate VARCHAR(20)	--開始的月份
DECLARE @Monthenddate VARCHAR(20)	--結束的月份
DECLARE @DateSet VARCHAR(20)		--計算月份用
DECLARE @RunDate VARCHAR(20)		--執行的月份

SET @StartStation = '{0}' 
SET @EndStation = '{1}'
SET @Openstardate = '{2}'
SET @Openenddate = '{3}'
SET @DateSet = CONCAT(DATEPART(YEAR,@Openstardate),'-',DATEPART(MONTH,@Openstardate) , '-01 0' , DATEPART(HOUR,@Openstardate) , ':0' , DATEPART(MINUTE,@Openstardate) , ':0' , DATEPART(SECOND,@Openstardate))
SET @Monthenddate = DATEADD(MONTH,1,@DateSet)
SET @Monthstardate = DATEADD(MONTH,-1,@Monthenddate)
SET @RunDate = @Monthstardate

SELECT TA.LocationId, ISNULL(E1.ECC_TxnAmt, 0) AS ECC_TxnAmt, ISNULL(E2.ECC_CSCAmt, 0) AS ECC_CSCAmt, ISNULL(E9.ECC_ExceptDeductAmt, 0) AS ECC_ExceptDeductAmt, ISNULL(E3.ECC_SaleAmt, 0) AS ECC_SaleAmt, ISNULL(ISNULL(E4.ECC_VAVMAddValueAmt, 0) - ISNULL(E5.ECC_VAVMCancelValueAmt, 0), 0) AS ECC_AddValueAmt, ISNULL(E6.ECC_PAMAddValueAmt, 0) AS ECC_PAMAddValueAmt, ISNULL(E7.ECC_PAMCancelValueAmt, 0) AS ECC_PAMCancelValueAmt, ISNULL(E8.ECC_AutoAddValueAmt, 0) AS ECC_AutoAddValueAmt,
    ISNULL(P1.IPASS_TxnAmt, 0) AS IPASS_TxnAmt, ISNULL(P2.IPASS_CSCAmt, 0) AS IPASS_CSCAmt, ISNULL(P9.IPASS_ExceptDeductAmt, 0) AS IPASS_ExceptDeductAmt, ISNULL(P3.IPASS_SaleAmt, 0) AS IPASS_SaleAmt, ISNULL(ISNULL(P4.IPASS_VAVMAddValueAmt, 0) - ISNULL(P5.IPASS_VAVMCancelValueAmt, 0), 0) AS IPASS_AddValueAmt, ISNULL(P6.IPASS_PAMAddValueAmt, 0) AS IPASS_PAMAddValueAmt, ISNULL(P7.IPASS_PAMCancelValueAmt, 0) AS IPASS_PAMCancelValueAmt, ISNULL(P8.IPASS_AutoAddValueAmt, 0) AS IPASS_AutoAddValueAmt,
    ISNULL(C1.ICASH_TxnAmt, 0) AS ICASH_TxnAmt, ISNULL(C2.ICASH_CSCAmt, 0) AS ICASH_CSCAmt, ISNULL(C9.ICASH_ExceptDeductAmt, 0) AS ICASH_ExceptDeductAmt, ISNULL(C3.ICASH_SaleAmt, 0) AS ICASH_SaleAmt, ISNULL(ISNULL(C4.ICASH_VAVMAddValueAmt, 0) - ISNULL(C5.ICASH_VAVMCancelValueAmt, 0), 0) AS ICASH_AddValueAmt, ISNULL(C6.ICASH_PAMAddValueAmt, 0) AS ICASH_PAMAddValueAmt, ISNULL(C7.ICASH_PAMCancelValueAmt, 0) AS ICASH_PAMCancelValueAmt, ISNULL(C8.ICASH_AutoAddValueAmt, 0) AS ICASH_AutoAddValueAmt,
    0 AS Credit_M_Exit, 0 AS Credit_M_Csc, 0 AS Credit_I_Exit, 0 AS Credit_I_Csc, 0 AS Credit_D_Exit, 0 AS Credit_D_Csc, 0 AS QrCode_P_Exit, 0 AS QrCode_P_Csc, 0 AS QrCode_T_Exit, 0 AS QrCode_T_Csc,
    ISNULL(T1.SOneTkt_Amt, 0) AS SOneTkt_Amt, ISNULL(T2.SOneTkt_Amt_discount, 0) AS SOneTkt_Amt_discount, ISNULL(T3.SOneTkt_bike_Amt, 0) AS SOneTkt_bike_Amt, ISNULL(T4.PTADuty, 0) AS PTADuty, ISNULL(T5.RefundSpeTkt, 0) AS RefundSpeTkt
FROM(
              SELECT LocationId  FROM Parm081_LocList WHERE CurrentType IN (1) AND LocationId BETWEEN @StartStation AND @EndStation) AS TA
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ECC_TxnAmt               FROM Txn_Exit              WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (22) AND IssueCode = 2                         GROUP BY LocationId) AS E1 ON TA.LocationId = E1.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ECC_CSCAmt               FROM Txn_ExcessByCSC       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (71) AND IssueCode = 2                         GROUP BY LocationId) AS E2 ON TA.LocationId = E2.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ECC_ExceptDeductAmt      FROM Txn_ExceptDeduct      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND IssueCode = 2 AND TxnType = 33                            GROUP BY LocationId) AS E9 ON TA.LocationId = E9.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ECC_SaleAmt              FROM Txn_SaleCard          WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (11) AND IssueCode = 2                         GROUP BY LocationId) AS E3 ON TA.LocationId = E3.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ECC_VAVMAddValueAmt      FROM Txn_AddValue          WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (1) AND IssueCode = 2 AND EquipType = 5        GROUP BY LocationId) AS E4 ON TA.LocationId = E4.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ECC_VAVMCancelValueAmt   FROM Txn_CancelValue       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (5) AND IssueCode = 2 AND EquipType = 5        GROUP BY LocationId) AS E5 ON TA.LocationId = E5.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ECC_PAMAddValueAmt       FROM Txn_AddValue          WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (1) AND IssueCode = 2 AND EquipType IN(3)      GROUP BY LocationId) AS E6 ON TA.LocationId = E6.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ECC_PAMCancelValueAmt    FROM Txn_CancelValue       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (5) AND IssueCode = 2 AND EquipType IN(3)      GROUP BY LocationId) AS E7 ON TA.LocationId = E7.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ECC_AutoAddValueAmt      FROM ECCTxn_Interface      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType = 2 AND MyTxnType = 2                             GROUP BY LocationId) AS E8 ON TA.LocationId = E8.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS IPASS_TxnAmt             FROM Txn_Exit              WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (22) AND IssueCode = 9                         GROUP BY LocationId) AS P1 ON TA.LocationId = P1.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS IPASS_CSCAmt             FROM Txn_ExcessByCSC       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (71) AND IssueCode = 9                         GROUP BY LocationId) AS P2 ON TA.LocationId = P2.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS IPASS_ExceptDeductAmt    FROM Txn_ExceptDeduct      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND IssueCode = 9 AND TxnType = 33                            GROUP BY LocationId) AS P9 ON TA.LocationId = P9.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS IPASS_SaleAmt            FROM Txn_SaleCard          WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (11) AND IssueCode = 9                         GROUP BY LocationId) AS P3 ON TA.LocationId = P3.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS IPASS_VAVMAddValueAmt    FROM Txn_AddValue          WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (1) AND IssueCode = 9 AND EquipType = 5        GROUP BY LocationId) AS P4 ON TA.LocationId = P4.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS IPASS_VAVMCancelValueAmt FROM Txn_CancelValue       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (1) AND IssueCode = 9 AND EquipType = 5        GROUP BY LocationId) AS P5 ON TA.LocationId = P5.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS IPASS_PAMAddValueAmt     FROM Txn_AddValue          WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (1) AND IssueCode = 9 AND EquipType IN(3)      GROUP BY LocationId) AS P6 ON TA.LocationId = P6.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS IPASS_PAMCancelValueAmt  FROM Txn_CancelValue       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (5) AND IssueCode = 9 AND EquipType IN(3)      GROUP BY LocationId) AS P7 ON TA.LocationId = P7.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS IPASS_AutoAddValueAmt    FROM ECCTxn_Interface      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType = 66                                              GROUP BY LocationId) AS P8 ON TA.LocationId = P8.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ICASH_TxnAmt             FROM Txn_Exit              WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (22) AND IssueCode = 11                        GROUP BY LocationId) AS C1 ON TA.LocationId = C1.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ICASH_CSCAmt             FROM Txn_ExcessByCSC       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (71) AND IssueCode = 11                        GROUP BY LocationId) AS C2 ON TA.LocationId = C2.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ICASH_ExceptDeductAmt    FROM Txn_ExceptDeduct      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND IssueCode = 11 AND TxnType = 33                           GROUP BY LocationId) AS C9 ON TA.LocationId = C9.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ICASH_SaleAmt            FROM Txn_SaleCard          WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (11) AND IssueCode = 11                        GROUP BY LocationId) AS C3 ON TA.LocationId = C3.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ICASH_VAVMAddValueAmt    FROM Txn_AddValue          WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (1) AND IssueCode = 11 AND EquipType = 5       GROUP BY LocationId) AS C4 ON TA.LocationId = C4.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ICASH_VAVMCancelValueAmt FROM Txn_CancelValue       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (5) AND IssueCode = 11 AND EquipType = 5       GROUP BY LocationId) AS C5 ON TA.LocationId = C5.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ICASH_PAMAddValueAmt     FROM Txn_AddValue          WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (1) AND IssueCode = 11 AND EquipType IN(3)     GROUP BY LocationId) AS C6 ON TA.LocationId = C6.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ICASH_PAMCancelValueAmt  FROM Txn_CancelValue       WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType in (5) AND IssueCode = 11 AND EquipType IN(3)     GROUP BY LocationId) AS C7 ON TA.LocationId = C7.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS ICASH_AutoAddValueAmt    FROM ECCTxn_Interface      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnType = 2 AND TxnSubType = 4                            GROUP BY LocationId) AS C8 ON TA.LocationId = C8.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS SOneTkt_Amt              FROM Txn_SellSpecialTicket WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND CardType in (1) AND EquipType Not IN (6)                  GROUP BY LocationId) AS T1 ON TA.LocationId = T1.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS SOneTkt_Amt_discount     FROM Txn_SellSpecialTicket WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND CardType in (3) AND EquipType Not IN (6)                  GROUP BY LocationId) AS T2 ON TA.LocationId = T2.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS SOneTkt_bike_Amt         FROM Txn_SellSpecialTicket WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND CardType in (2) AND EquipType Not IN (6)                  GROUP BY LocationId) AS T3 ON TA.LocationId = T3.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS PTADuty                  FROM Txn_SellSpecialTicket WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND EquipType IN (6)                                          GROUP BY LocationId) AS T4 ON TA.LocationId = T4.LocationId
    LEFT JOIN(SELECT LocationId, SUM(TxnAmt) AS RefundSpeTkt             FROM Txn_RefundSpeTkt      WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate                                                               GROUP BY LocationId) AS T5 ON TA.LocationId = T5.LocationId
ORDER BY LocationId ";
        public static readonly string Sql_Command_Month_EquipAmount_Day = @"--設備營收月報表(天期)
DECLARE @StartStation VARCHAR(10)	--該路線車站起點
DECLARE @EndStation VARCHAR(10)		--該路線車站終點
DECLARE @Openstardate VARCHAR(20)	--搜尋時間開始
DECLARE @Openenddate VARCHAR(20)	--搜尋時間結束
DECLARE @Monthstardate VARCHAR(20)	--開始的月份
DECLARE @Monthenddate VARCHAR(20)	--結束的月份
DECLARE @DateSet VARCHAR(20)		--計算月份用
DECLARE @RunDate VARCHAR(20)		--執行的月份

SET @StartStation = '{0}' 
SET @EndStation = '{1}'
SET @Openstardate = '{2}'
SET @Openenddate = '{3}'
SET @DateSet = CONCAT(DATEPART(YEAR,@Openstardate),'-',DATEPART(MONTH,@Openstardate) , '-01 0' , DATEPART(HOUR,@Openstardate) , ':0' , DATEPART(MINUTE,@Openstardate) , ':0' , DATEPART(SECOND,@Openstardate))
SET @Monthenddate = DATEADD(MONTH,1,@DateSet)
SET @Monthstardate = DATEADD(MONTH,-1,@Monthenddate)
SET @RunDate = @Monthstardate

SELECT LocationId as mLoc
INTO #mTempLocation
FROM Parm081_LocList
WHERE LocationId >= @StartStation AND LocationId <= @EndStation AND CurrentType IN (1)

SELECT *
INTO #mTempTicket
FROM Txn_SellSpecialTicket JOIN #mTempLocation on LocationId = mLoc
WHERE TxnDT >= @Monthstardate AND TxnDT < @Monthenddate

SELECT *
INTO #mTempReTicket
FROM Txn_RefundSpeTkt JOIN #mTempLocation on LocationId = mLoc
WHERE TxnDT >= @Monthstardate AND TxnDT < @Monthenddate

CREATE TABLE #mTemplate (
    mYear INT,
    mMonth INT,
    mDay INT,
    Ecc_ExitAmt             INT,    --悠遊卡_出站
    Ecc_Csc                 INT,    --悠遊卡_罰款
    Ecc_Dec                 INT,    --悠遊卡_異常處理
    Ecc_VAVM_Sale           INT,    --悠遊卡_售卡
    Ecc_VAVM_AddValue       INT,    --悠遊卡_VAVM加值
    Ecc_PAM_AddValue        INT,    --悠遊卡_PAM加值
    Ecc_Cancel              INT,    --悠遊卡_取消加值
    Ecc_AutoValue           INT,    --悠遊卡_自動加值
    
    Ipass_ExitAmt           INT,    --一卡通_出站
    Ipass_Csc               INT,    --一卡通_罰款
    Ipass_Dec               INT,    --一卡通_異常處理
    Ipass_VAVM_Sale         INT,    --一卡通_售卡
    Ipass_VAVM_AddValue     INT,    --一卡通_VAVM加值
    Ipass_PAM_AddValue      INT,    --一卡通_PAM加值
    Ipass_Cancel            INT,    --一卡通_取消加值
    Ipass_AutoValue         INT,    --一卡通_自動加值
    
    Icash_ExitAmt           INT,    --愛金卡_出站
    Icash_Csc               INT,    --愛金卡_罰款
    Icash_Dec               INT,    --愛金卡_異常處理
    Icash_VAVM_Sale         INT,    --愛金卡_售卡
    Icash_VAVM_AddValue     INT,    --愛金卡_VAVM加值
    Icash_PAM_AddValue      INT,    --愛金卡_PAM加值
    Icash_Cancel            INT,    --愛金卡_取消加值
    Icash_AutoValue         INT,    --愛金卡_自動加值

    Credit_M_Exit           INT,    --信用卡_萬事達_出站
    Credit_M_Csc            INT,    --信用卡_萬事達_罰款
    Credit_I_Exit           INT,    --信用卡_銀聯_出站
    Credit_I_Csc            INT,    --信用卡_銀聯_罰款
    Credit_D_Exit           INT,    --信用卡_發現卡_出站
    Credit_D_Csc            INT,    --信用卡_發現卡_罰款
    QrCode_P_Exit           INT,    --乘車碼_一卡通_出站
    QrCode_P_Csc            INT,    --乘車碼_一卡通_罰款
    QrCode_T_Exit           INT,    --乘車碼_台新_出站
    QrCode_T_Csc            INT,    --乘車碼_台新_罰款

    VAVM_OneTime_General    INT,    --系統售一般單程票
    VAVM_OneTime_Concession INT,    --系統售優待單程票
    VAVM_OneTime_Bike       INT,    --系統售自行車單程票
    PTA_ReTicket            INT,    --查票機補票
    PAM_BackTicket          INT        --系統退票
)

WHILE(@RunDate < @Monthenddate)
BEGIN
    INSERT INTO #mTemplate     
    SELECT 
          DATEPART(YEAR,@RunDate) AS mYear, DATEPART(MONTH,@RunDate) AS mMonth, DATEPART(DAY,@RunDate) AS mDay
        , ISNULL(Ecc_ExitAmt              ,0) AS Ecc_ExitAmt            
        , ISNULL(Ecc_Csc                  ,0) AS Ecc_Csc                
        , ISNULL(Ecc_Dec                  ,0) AS Ecc_Dec                
        , ISNULL(Ecc_VAVM_Sale            ,0) AS Ecc_VAVM_Sale          
        , ISNULL(Ecc_VAVM_AddValue        ,0) AS Ecc_VAVM_AddValue      
        , ISNULL(Ecc_PAM_AddValue         ,0) AS Ecc_PAM_AddValue       
        , ISNULL(Ecc_Cancel               ,0) AS Ecc_Cancel             
        , ISNULL(Ecc_AutoValue            ,0) AS Ecc_AutoValue          

        , ISNULL(Ipass_ExitAmt            ,0) AS Ipass_ExitAmt          
        , ISNULL(Ipass_Csc                ,0) AS Ipass_Csc              
        , ISNULL(Ipass_Dec                ,0) AS Ipass_Dec              
        , ISNULL(Ipass_VAVM_Sale          ,0) AS Ipass_VAVM_Sale        
        , ISNULL(Ipass_VAVM_AddValue      ,0) AS Ipass_VAVM_AddValue    
        , ISNULL(Ipass_PAM_AddValue       ,0) AS Ipass_PAM_AddValue     
        , ISNULL(Ipass_Cancel             ,0) AS Ipass_Cancel           
        , ISNULL(Ipass_AutoValue          ,0) AS Ipass_AutoValue        
 
        , ISNULL(Icash_ExitAmt            ,0) AS Icash_ExitAmt          
        , ISNULL(Icash_Csc                ,0) AS Icash_Csc              
        , ISNULL(Icash_Dec                ,0) AS Icash_Dec              
        , ISNULL(Icash_VAVM_Sale          ,0) AS Icash_VAVM_Sale        
        , ISNULL(Icash_VAVM_AddValue      ,0) AS Icash_VAVM_AddValue    
        , ISNULL(Icash_PAM_AddValue       ,0) AS Icash_PAM_AddValue     
        , ISNULL(Icash_Cancel             ,0) AS Icash_Cancel           
        , ISNULL(Icash_AutoValue          ,0) AS Icash_AutoValue        

        , ISNULL(0                        ,0) AS Credit_M_Exit          
        , ISNULL(0                        ,0) AS Credit_M_Csc           
        , ISNULL(0                        ,0) AS Credit_I_Exit          
        , ISNULL(0                        ,0) AS Credit_I_Csc           
        , ISNULL(0                        ,0) AS Credit_D_Exit          
        , ISNULL(0                        ,0) AS Credit_D_Csc           
        , ISNULL(0                        ,0) AS QrCode_P_Exit          
        , ISNULL(0                        ,0) AS QrCode_P_Csc           
        , ISNULL(0                        ,0) AS QrCode_T_Exit          
        , ISNULL(0                        ,0) AS QrCode_T_Csc           

        , ISNULL(VAVM_OneTime_General     ,0) AS VAVM_OneTime_General   
        , ISNULL(VAVM_OneTime_Concession  ,0) AS VAVM_OneTime_Concession
        , ISNULL(VAVM_OneTime_Bike        ,0) AS VAVM_OneTime_Bike      
        , ISNULL(PTA_ReTicket             ,0) AS PTA_ReTicket           
        , ISNULL(PAM_BackTicket           ,0) AS PAM_BackTicket         

    From
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Ecc_ExitAmt                FROM Txn_Exit           JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (2))                              AS Ecc_ExitAmt                      LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Ecc_Csc                    FROM Txn_ExcessByCSC    JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (2))                              AS Ecc_Csc                  ON 1=1  LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Ecc_Dec                    FROM Txn_ExceptDeduct   JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (2))                              AS Ecc_Dec                  ON 1=1  LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Ecc_VAVM_Sale              FROM Txn_SaleCard       JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (2))                              AS Ecc_VAVM_Sale            ON 1=1  LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Ecc_VAVM_AddValue          FROM Txn_AddValue       JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (2) AND EquipType IN (5))         AS Ecc_VAVM_AddValue        ON 1=1  LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Ecc_PAM_AddValue           FROM Txn_AddValue       JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (2) AND EquipType IN (3))         AS Ecc_PAM_AddValue         ON 1=1  LEFT JOIN     
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Ecc_Cancel                 FROM Txn_CancelValue    JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (2))                              AS Ecc_Cancle               ON 1=1  LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Ecc_AutoValue              FROM Txn_Autoload       JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (2))                              AS Ecc_AutoValue            ON 1=1  LEFT JOIN 

        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Ipass_ExitAmt              FROM Txn_Exit           JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (9))                              AS Ipass_ExitAmt            ON 1=1  LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Ipass_Csc                  FROM Txn_ExcessByCSC    JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (9))                              AS Ipass_Csc                ON 1=1  LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Ipass_Dec                  FROM Txn_ExceptDeduct   JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (9))                              AS Ipass_Dec                ON 1=1  LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Ipass_VAVM_Sale            FROM Txn_SaleCard       JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (9))                              AS Ipass_VAVM_Sale          ON 1=1  LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Ipass_VAVM_AddValue        FROM Txn_AddValue       JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (9) AND EquipType IN (5))         AS Ipass_VAVM_AddValue      ON 1=1  LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Ipass_PAM_AddValue         FROM Txn_AddValue       JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (9) AND EquipType IN (3))         AS Ipass_PAM_AddValue       ON 1=1  LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Ipass_Cancel               FROM Txn_CancelValue    JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (9))                              AS Ipass_Cancle             ON 1=1  LEFT JOIN     
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Ipass_AutoValue            FROM Txn_Autoload       JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (9))                              AS Ipass_AutoValue          ON 1=1  LEFT JOIN 

        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Icash_ExitAmt              FROM Txn_Exit           JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (11))                             AS Icash_ExitAmt            ON 1=1  LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Icash_Csc                  FROM Txn_ExcessByCSC    JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (11))                             AS Icash_Csc                ON 1=1  LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Icash_Dec                  FROM Txn_ExceptDeduct   JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (11))                             AS Icash_Dec                ON 1=1  LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Icash_VAVM_Sale            FROM Txn_SaleCard       JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (11))                             AS Icash_VAVM_Sale          ON 1=1  LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Icash_VAVM_AddValue        FROM Txn_AddValue       JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (11) AND EquipType IN (5))        AS Icash_VAVM_AddValue      ON 1=1  LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Icash_PAM_AddValue         FROM Txn_AddValue       JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (11) AND EquipType IN (3))        AS Icash_PAM_AddValue       ON 1=1  LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Icash_Cancel               FROM Txn_CancelValue    JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (11))                             AS Icash_Cancle             ON 1=1  LEFT JOIN     
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS Icash_AutoValue            FROM Txn_Autoload       JOIN #mTempLocation ON LocationId = mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND IssueCode IN (11))                             AS Icash_AutoValue          ON 1=1  LEFT JOIN 

        (SELECT ISNULL(0              ,0) AS Credit_M_Exit)                                                                                                                                                                                                                                                 AS Credit_M_Exit            ON 1=1    LEFT JOIN 
        (SELECT ISNULL(0              ,0) AS Credit_M_Csc)                                                                                                                                                                                                                                                  AS Credit_M_Csc             ON 1=1    LEFT JOIN     
        (SELECT ISNULL(0              ,0) AS Credit_I_Exit)                                                                                                                                                                                                                                                 AS Credit_I_Exit            ON 1=1    LEFT JOIN 
        (SELECT ISNULL(0              ,0) AS Credit_I_Csc)                                                                                                                                                                                                                                                  AS Credit_I_Csc             ON 1=1    LEFT JOIN     
        (SELECT ISNULL(0              ,0) AS Credit_D_Exit)                                                                                                                                                                                                                                                 AS Credit_D_Exit            ON 1=1    LEFT JOIN 
        (SELECT ISNULL(0              ,0) AS Credit_D_Csc)                                                                                                                                                                                                                                                  AS Credit_D_Csc             ON 1=1    LEFT JOIN     
        (SELECT ISNULL(0              ,0) AS QrCode_P_Exit)                                                                                                                                                                                                                                                 AS QrCode_P_Exit            ON 1=1    LEFT JOIN 
        (SELECT ISNULL(0              ,0) AS QrCode_P_Csc)                                                                                                                                                                                                                                                  AS QrCode_P_Csc             ON 1=1    LEFT JOIN     
        (SELECT ISNULL(0              ,0) AS QrCode_T_Exit)                                                                                                                                                                                                                                                 AS QrCode_T_Exit            ON 1=1    LEFT JOIN 
        (SELECT ISNULL(0              ,0) AS QrCode_T_Csc)                                                                                                                                                                                                                                                  AS QrCode_T_Csc             ON 1=1    LEFT JOIN     

        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS VAVM_OneTime_General       FROM #mTempTicket       JOIN #mTempLocation AS mL ON LocationId = mL.mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND EquipType IN (5) AND CardType IN (1)) AS VAVM_OneTime_General     ON 1=1    LEFT JOIN     
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS VAVM_OneTime_Concession    FROM #mTempTicket       JOIN #mTempLocation AS mL ON LocationId = mL.mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND EquipType IN (5) AND CardType IN (3)) AS VAVM_OneTime_Concession  ON 1=1    LEFT JOIN     
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS VAVM_OneTime_Bike          FROM #mTempTicket       JOIN #mTempLocation AS mL ON LocationId = mL.mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND EquipType IN (5) AND CardType IN (2)) AS VAVM_OneTime_Bike        ON 1=1    LEFT JOIN 
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS PTA_ReTicket               FROM #mTempTicket       JOIN #mTempLocation AS mL ON LocationId = mL.mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate) AND EquipType NOT IN (5))                 AS PTA_ReTicket             ON 1=1    LEFT JOIN     
        (SELECT ISNULL(SUM(TxnAmt)    ,0) AS PAM_BackTicket             FROM #mTempReTicket     JOIN #mTempLocation AS mL ON LocationId = mL.mLoc WHERE TxnDT >= @Openstardate AND TxnDT < @Openenddate AND TxnDT >= @RunDate AND TxnDt < DATEADD(day,1,@RunDate))                                          AS PAM_BackTicket           ON 1=1

    SET @RunDate = DATEADD(day,1,@RunDate)
END

select *
from #mTemplate
order by mYear,mMonth,mDay

DROP TABLE #mTempReTicket
DROP TABLE #mTempTicket
DROP TABLE #mTempLocation
DROP TABLE #mTemplate

";
        public static readonly string Sql_Command_19 = @"";
        public static readonly string Sql_Command_20 = @"";

        #endregion

    }

    public class SubLoading
    {
        private MyCommonds MyCommond = new MyCommonds();
        private string ThisReceive = "SubLoading";
        Boolean Status_Config_Updata_mContorl = true;   // 是否啟用更新
        LrtMain MainLrt;

        public SubLoading(LrtMain MainServer)
        {
            MyCommond.WriteLog(ThisReceive, $"創建新變數");
            MainLrt = MainServer;

            MyCommond.CheckFolder(MyCreat.Path_Sql + @"./Old/");
        }

        public void Test()
        {

        }

        public void Loading_Normal()
        {
            string path = $"{MyCommond.Path_Config}/{MainLrt.Config_Setting}";
            CheckConfig(path, MyCreat.Lrt_Config_Txt, MyCreat.Lrt_UpData_Config_Txt);
            string[] File_ProgramConfig = File.ReadAllLines(path, Encoding.UTF8);
            //SqlServer sqlServer = new SqlServer();
            foreach (string item in File_ProgramConfig)
            {
                //this.WriteLog(ExecutionMode.Simple, item);
                string[] FileLine = item.Split(',');
                switch (FileLine[0])
                {
                    case "Using Mode": /*           */ MainLrt.Using_Mode /*              */ = checkMode(FileLine[1]); /*               */ break;
                    case "Export Language": /*      */ MainLrt.MyCommond.ExportLanguage /**/ = checkLanguage(FileLine[1]); /*           */ break;
                    //case "Excel Applaction Show": /**/ MainLrt.Excel_Applaction_Show /*   */ = FileLine[1] == "1"; /*                   */ break;
                    //case "Install MobilePay": /*    */ MainLrt.Install_MobilePay /*       */ = FileLine[1] == "1"; /*                   */ break;
                    //case "Auto Pao Excel": /*       */ MainLrt.Auto_Pao_Excel /*          */ = FileLine[1] == "1"; /*                   */ break;
                    //case "Operation First Day": /*  */ MainLrt.Operation_First_Day /*     */ = Convert.ToDateTime(FileLine[1]); /*      */ break;
                    //case "Using Line": /*           */ MainLrt.OperationLine /*           */ = FileLine[1].Trim(); /*                   */ break;
                    //case "Test Button": /*          */ MainLrt.Test_Button /*             */ = FileLine[1] == "1"; /*                   */ break;
                    //case "Finish Button": /*        */ MainLrt.Finish_Button /*           */ = Convert.ToInt16(FileLine[1].Trim()); /*  */ break;
                    //case "Sql Server Ip": /*        */ MainLrt.SqlServers.IP /*            */ = Convert.ToString(FileLine[1].Trim()); /**/ break;
                    //case "Sql Catalog": /*          */ MainLrt.SqlServers.Calalog /*       */ = Convert.ToString(FileLine[1].Trim()); /**/ break;
                    //case "Sql User Id": /*          */ MainLrt.SqlServers.ID /*            */ = Convert.ToString(FileLine[1].Trim()); /**/ break;
                    //case "Sql User Password": /*    */ MainLrt.SqlServers.Password /*      */ = Convert.ToString(FileLine[1].Trim()); /**/ break;
                    default: break;
                }
            }
            //temp.SqlServers.Add(sqlServer);
            
            Loading_Line();
        }
        
        private void Loading_Line()
        {
            string path = MyCommond.Path_Config + MainLrt.Config_OperationLine;
            CheckConfig(path,MyCreat.Lrt_Config_Line,MyCreat.Lrt_UpData_Config_Line);
            #region 載入路線設定
            MyCommond.WriteLog(ThisReceive, $"載入路線設定");
            string[] File_ProgramConfig = File.ReadAllLines(path,Encoding.UTF8);
            Boolean HaveOperationLine = false;
            foreach (string line in File_ProgramConfig)
            {
                MyCommond.WriteLog(ThisReceive, $"讀取：{line}");
                string[] FileLine = line.Split(new char[2] { ',', '\t' });
                //if (FileLine[1].Trim() == MainLrt.OperationLine)
                //{
                //    MyCommond.WriteLog(ThisReceive, $"有與營運路線設定檔相同路線名");
                //    MainLrt.OperationLine_String = Convert.ToString(FileLine[2].Trim());
                //    HaveOperationLine = true;
                //    break;
                //}
            }
            File_ProgramConfig = null;
            #endregion

        }

        public void Loading_Operate()
        {
            //MainLrt.Stock_Lrt_ReportFile2 /* */ = Load2Model<ReportFileTemp2> /* */ ($"{MyCommond.Path_Config}/{MainLrt.Config_Report}" /*      */, MyCreat.UpData_Lrt_Report_Csv /*      */, MyCreat.Lrt_Report_Csv /*   */, (new ReportFileTemp2()).ExportTitle_zhTW() /* */, 1);
            //MainLrt.Stock_MobileCardType /*  */ = Load2Model<MobileCardType> /*  */ ($"{MyCommond.Path_Config}/{MainLrt.Config_MobileCard}" /*  */, MyCreat.UpData_MobileCard_Csv /*      */, MyCreat.MobileCard_Csv /*   */, (new MobileCardType()).ExportTitle_zhTW() /*  */, 1);
            
            //var item /*                      */ = Load2Model<LrtStationList> /*  */ ($"{MyCommond.Path_Config}/{MainLrt.Config_StationList}" /* */, MyCreat.UpData_Lrt_StationList_Csv /* */, MyCreat.Lrt_Station_Csv /*  */, (new LrtStationList()).ExportTitle_zhTW() /*  */, 1);
            //MainLrt.Stock_Lrt_Station = item.FindAll(x => x.OperationLine == MainLrt.OperationLine).ToList();
            
            //var jtem /*                      */ = Load2Model<Subsidy> /*         */ ($"{MyCommond.Path_Config}/{MainLrt.Config_Subidy}" /*      */, MyCreat.UpData_Lrt_Subsidy_Csv /*     */, MyCreat.Lrt_Subsidy_Csv /*  */, (new Subsidy()).ExportTitle_zhTW() /*         */, 1);
            //MainLrt.Stock_Lrt_Subsidy = jtem.FindAll(x => x.Operation_Area == MainLrt.OperationLine).ToList();
        }

        public void Loading_Sql()
        {
            //#region SQL語法設定

            //MyCommond.WriteLog(ThisReceive, $"載入SQL語法設定檔設定");
            //CheckConfig(MyCommond.Path_Config + MainLrt.Config_SqlList, MyCreat.Lrt_Config_Sql_List, MyCreat.UpData_Config_Sql_List);
            //string[] File_ProgramConfig = File.ReadAllLines(MyCommond.Path_Config + MainLrt.Config_SqlList, Encoding.Default);
            //bool JumpFirstLine = true;
            //MainLrt.Stock_SqlList = new List<SqlList>();
            ////msgstr = @"資料表中文:{0}, 檔案名稱1:{1}, 檔案名稱2:{2}, 備註:{3}";
            //foreach (string line in File_ProgramConfig)
            //{
            //    // WriteLog(1, MethodBase.GetCurrentMethod().Name.ToString(), "讀取：" + line);
            //    string[] FileLine = line.Split(new char[2] { ',', '\t' });

            //    if (!JumpFirstLine)
            //    {
            //        //WriteLog(1, MethodBase.GetCurrentMethod().Name.ToString(), string.Format(msgstr, FileLine[0], FileLine[1], FileLine[2], FileLine[3]));
            //        // WriteLog(1, MethodBase.GetCurrentMethod().Name.ToString(), $"資料表中文:{FileLine[0]}, 檔案名稱1:{FileLine[1]}, 檔案名稱2:{FileLine[2]}, 備註:{FileLine[3]}");
            //        MainLrt.Stock_SqlList.Add(new SqlList()
            //        {
            //            UseNameCh = Convert.ToString(FileLine[0]),
            //            FileName1 = Convert.ToString(FileLine[1]),
            //            FileName2 = Convert.ToString(FileLine[2]),
            //            Remark = Convert.ToString(FileLine[3])
            //        });
            //    }
            //    else
            //    {
            //        JumpFirstLine = false;
            //    }

            //}
            //File_ProgramConfig = null;

            //MyCommond.WriteLog(ThisReceive, $"載入SQL語法");
            //foreach (SqlList SqlLine in MainLrt.Stock_SqlList)
            //{
            //    MyCommond.WriteLog(ThisReceive, $"讀取：{SqlLine.ToString()}");
            //    string str = SqlLine.UseNameCh;
            //    if (SqlLine.FileName1 == "-" || SqlLine.Remark == "停用") { }
            //    else
            //    {
            //        LrtTxnClientSwitch LTCS = SearchSql(str);
            //        if (LTCS == LrtTxnClientSwitch.NA) continue;
            //        SqlLine.Using = ReadSqlCommond(MyCreat.Path_Sql + SqlLine.FileName1 + ".txt", LTCS);
            //    }
            //}

            //#endregion
        }
        
        private LrtTxnClientSwitch SearchSql(string SqlName)
        {
            LrtTxnClientSwitch reSql = LrtTxnClientSwitch.NA;
                 if (SqlName == LrtTxnClientSwitch.Day_Volume.ToDescription())                     {reSql = LrtTxnClientSwitch.Day_Volume                     ;}
            else if (SqlName == LrtTxnClientSwitch.Day_Amount.ToDescription())                     {reSql = LrtTxnClientSwitch.Day_Amount                     ;}
            else if (SqlName == LrtTxnClientSwitch.Day_EachStation_EachTime.ToDescription())       {reSql = LrtTxnClientSwitch.Day_EachStation_EachTime       ;}
            else if (SqlName == LrtTxnClientSwitch.Day_ElectronicTicket.ToDescription())           {reSql = LrtTxnClientSwitch.Day_ElectronicTicket           ;}
            else if (SqlName == LrtTxnClientSwitch.Day_AllRideList.ToDescription())                {reSql = LrtTxnClientSwitch.Day_AllRideList                ;}
            else if (SqlName == LrtTxnClientSwitch.Day_TrafficAmount.ToDescription())              {reSql = LrtTxnClientSwitch.Day_TrafficAmount              ;}
            else if (SqlName == LrtTxnClientSwitch.Day_OriginDestination.ToDescription())          {reSql = LrtTxnClientSwitch.Day_OriginDestination          ;}
            else if (SqlName == LrtTxnClientSwitch.Day_EquipAmount.ToDescription())                {reSql = LrtTxnClientSwitch.Day_EquipAmount                ;}
            else if (SqlName == LrtTxnClientSwitch.Month_AllRideList.ToDescription())              {reSql = LrtTxnClientSwitch.Month_AllRideList              ;}
            else if (SqlName == LrtTxnClientSwitch.Month_OriginDestination.ToDescription())        {reSql = LrtTxnClientSwitch.Month_OriginDestination        ;}
            else if (SqlName == LrtTxnClientSwitch.Month_ElectronicTicket_Station.ToDescription()) {reSql = LrtTxnClientSwitch.Month_ElectronicTicket_Station ;}
            else if (SqlName == LrtTxnClientSwitch.Month_ElectronicTicket_Day.ToDescription())     {reSql = LrtTxnClientSwitch.Month_ElectronicTicket_Day     ;}
            else if (SqlName == LrtTxnClientSwitch.Month_OwnTicketVolume_Station.ToDescription())  {reSql = LrtTxnClientSwitch.Month_OwnTicketVolume_Station  ;}
            else if (SqlName == LrtTxnClientSwitch.Month_OwnTicketVolume_Day.ToDescription())      {reSql = LrtTxnClientSwitch.Month_OwnTicketVolume_Day      ;}
            else if (SqlName == LrtTxnClientSwitch.Month_TrafficAmount_Station.ToDescription())    {reSql = LrtTxnClientSwitch.Month_TrafficAmount_Station    ;}
            else if (SqlName == LrtTxnClientSwitch.Month_TrafficAmount_Day.ToDescription())        {reSql = LrtTxnClientSwitch.Month_TrafficAmount_Day        ;}
            else if (SqlName == LrtTxnClientSwitch.Month_EquipAmount_Station.ToDescription())      {reSql = LrtTxnClientSwitch.Month_EquipAmount_Station      ;}
            else if (SqlName == LrtTxnClientSwitch.Month_EquipAmount_Day.ToDescription())          {reSql = LrtTxnClientSwitch.Month_EquipAmount_Day          ; }
            return reSql;
        }

        private ReSql SearchSql(LrtTxnClientSwitch methonName)
        {
            ReSql reSql = new ReSql();
                 if (methonName == LrtTxnClientSwitch.Day_Volume)                     {reSql.Using = MyCreat.Sql_Command_Day_Volume; reSql.UpData = true;}
            else if (methonName == LrtTxnClientSwitch.Day_Amount)                     {reSql.Using = MyCreat.Sql_Command_Day_Amount; reSql.UpData = true;}
            else if (methonName == LrtTxnClientSwitch.Day_EachStation_EachTime)       {reSql.Using = MyCreat.Sql_Command_Day_EachStation_EachTime; reSql.UpData = true;}
            else if (methonName == LrtTxnClientSwitch.Day_ElectronicTicket)           {reSql.Using = MyCreat.Sql_Command_Day_ElectronicTicket; reSql.UpData = true;}
            else if (methonName == LrtTxnClientSwitch.Day_AllRideList)                {reSql.Using = MyCreat.Sql_Command_Day_AllRideList; reSql.UpData = true;}
            else if (methonName == LrtTxnClientSwitch.Day_TrafficAmount)              {reSql.Using = MyCreat.Sql_Command_Day_TrafficAmount; reSql.UpData = true;}
            else if (methonName == LrtTxnClientSwitch.Day_OriginDestination)          {reSql.Using = MyCreat.Sql_Command_Day_OriginDestination; reSql.UpData = true;}
            else if (methonName == LrtTxnClientSwitch.Day_EquipAmount)                {reSql.Using = MyCreat.Sql_Command_Day_EquipAmount; reSql.UpData = true;}
            else if (methonName == LrtTxnClientSwitch.Month_AllRideList)              {reSql.Using = MyCreat.Sql_Command_Month_AllRideList; reSql.UpData = true;}
            else if (methonName == LrtTxnClientSwitch.Month_OriginDestination)        {reSql.Using = MyCreat.Sql_Command_Month_OriginDestination; reSql.UpData = true;}
            else if (methonName == LrtTxnClientSwitch.Month_ElectronicTicket_Station) {reSql.Using = MyCreat.Sql_Command_Month_ElectronicTicket_Station; reSql.UpData = true;}
            else if (methonName == LrtTxnClientSwitch.Month_ElectronicTicket_Day)     {reSql.Using = MyCreat.Sql_Command_Month_ElectronicTicket_Day; reSql.UpData = true;}
            else if (methonName == LrtTxnClientSwitch.Month_OwnTicketVolume_Station)  {reSql.Using = MyCreat.Sql_Command_Month_OwnTicketVolume_Station; reSql.UpData = true;}
            else if (methonName == LrtTxnClientSwitch.Month_OwnTicketVolume_Day)      {reSql.Using = MyCreat.Sql_Command_Month_OwnTicketVolume_Day; reSql.UpData = true;}
            else if (methonName == LrtTxnClientSwitch.Month_TrafficAmount_Station)    {reSql.Using = MyCreat.Sql_Command_Month_TrafficAmount_Station; reSql.UpData = true;}
            else if (methonName == LrtTxnClientSwitch.Month_TrafficAmount_Day)        {reSql.Using = MyCreat.Sql_Command_Month_TrafficAmount_Day; reSql.UpData = true;}
            else if (methonName == LrtTxnClientSwitch.Month_EquipAmount_Station)      {reSql.Using = MyCreat.Sql_Command_Month_EquipAmount_Station; reSql.UpData = true;}
            else if (methonName == LrtTxnClientSwitch.Month_EquipAmount_Day)          {reSql.Using = MyCreat.Sql_Command_Month_EquipAmount_Day; reSql.UpData = true;}
            return reSql;
        }

        private string ReadSqlCommond(string path, LrtTxnClientSwitch MethodName)
        {
            ReSql reSql = SearchSql(MethodName);
            CheckConfig(path, reSql.Using, reSql.UpData);
            return File.ReadAllText(path, Encoding.Default);
        }

        public List<PaoTicketSaleList> PaoStation()
        {
            //List<PaoTicketSaleList> item = Load2Model<PaoTicketSaleList>(
            //    MyCommond.Path_Config + MainLrt.Config_PaoList,
            //    MyCreat.UpData_Lrt_Pao_Csv,
            //    MyCreat.Lrt_Pao_Csv,
            //    (new PaoTicketSaleList()).ExportTitle_zhTW(),
            //    1);
            //var jtem = item.FindAll(x => x.Operation == MainLrt.OperationLine).ToList();

            //return jtem;
            return null;
        }

        private ExecutionMode checkMode(string str)
        {
            ExecutionMode mode = ExecutionMode.Normal;

            switch (str)
            {
                case "Simple": mode = ExecutionMode.Simple; break;
                case "Normal": mode = ExecutionMode.Normal; break;
                case "Debug": mode = ExecutionMode.Debug; break;
            }

            return mode;
        }

        private Language checkLanguage(string str)
        {
            Language mode = Language.zhTW;

            switch (str)
            {
                case "zhTW":
                case "Taiwan":
                case "Taiwanese": mode = Language.zhTW; break;
                case "Number": //mode = Language.Number; break;
                case "zhCN":
                case "Chinese":
                case "China": //mode = Language.zhCN; break;
                case "Eng":
                case "English":
                default: mode = Language.English; break;
            }

            return mode;
        }

        private void CheckConfig(string Path, string ConfigData, Boolean UpData)
        {
            var sss = Path.Split('/');
            MyCommond.WriteLog(ThisReceive, $"檢查檔案 {sss[sss.Length - 1]}");
            DateTime TmPg = Convert.ToDateTime(File.GetLastWriteTime(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName)); // 程式修改時間
            DateTime TmCf = Convert.ToDateTime(Directory.GetCreationTime(Path)); // 檔案創建時間

            if (Status_Config_Updata_mContorl && UpData && TmPg > TmCf)
            {
                if ( MyCommond.CheckFile(Path)) File.Delete(Path);
                if (!MyCommond.CheckFile(Path)) File.WriteAllText(Path, ConfigData, Encoding.UTF8);
            }
        }

        private List<TResult> Load2Model<TResult>(string path, Boolean UpData, string Data, string TResultTile, int DataRow = 0) where TResult : class, new()
        {
            var sss = path.Split('/');
            CheckConfig(path, Data, UpData);
            MyCommond.WriteLog(ThisReceive, $"載入檔案{sss[sss.Length - 1]}");
            var item = MyCommond.ToViewModel<TResult>(ThisReceive, path, TResultTile, DataRow);
            return item;
        }

        private List<TResult> LoadSubsidy<TResult>(string path, Boolean UpData, string Data, int DataRow = 0) where TResult : class, new()
        {
            CheckConfig(path, Data, UpData);
            MyCommond.WriteLog(ThisReceive, $"載入票差補貼檔設定");
            List<TResult> MainLrtList = new List<TResult>();
            Type type = typeof(TResult);
            string[] File_ProgramConfig = File.ReadAllLines(path, Encoding.Default);
            string msgstr = @"進站車站:{0}, 出站車站:{1}, 補貼金額:{2}";

            TResult tr = new TResult();
            for (int i = DataRow; i < File_ProgramConfig.Length; i++)
            {
                string[] FileSplit = File_ProgramConfig[i].Split(',');
                for (int j = 1; j < FileSplit.Length; j++)
                {
                    type.GetProperties()[j].SetValue(tr, Convert.ToString(FileSplit[j]), null);
                    //string ss = MainLrt.Stock_Lrt_Station[i - 1].CodeNumber.ToString();
                }
            }

            MainLrtList.Add(tr);
            TResult MainLrtData = new TResult();



            return MainLrtList;
        }

    }
}