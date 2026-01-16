using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//--
using System.Data.SQLite;
using System.Diagnostics;
using System.IO;
using UsuallyCommond.MyEnum;
using System.Windows.Forms;

namespace UsuallyCommond
{
    public class MySQLite
    {
        private MyCommonds MyCommond;
        private string ThisReceive;
        public int MaxCount { get; private set; }
        public int MinCount { get; private set; }
        public int CountValue { get; private set; }

        public string DB_Title { get; private set; }
        public string DB_Path { get; private set; }
        public string DB_FullPath { get; private set; }
        public string DB_Line { get; set; }
        public string DB_Line_Title { get; private set; }
        public string DB_Name { get; private set; }
        public string DB_FullName { get; private set; }
        public string DB_Connect { get; private set; }
        public StatusEnum Status { get; private set; }
        public DateTime DtStart { get; set; }
        public DateTime DtEnd { get; set; }
        public DateTime ReadDt { get; set; }
        public string SetCommand { get; set; }
        public Boolean Stop { get; set; }
        public int LoadLimit { get; set; }

        public MySQLite()
        {
            MyCommond = new MyCommonds();
            ThisReceive = "MySQLite";

            DB_Title = "Data Source=";
            DB_Path = @"DataBase\";

            DateTime dt = Convert.ToDateTime(DateTime.Now.ToString("yyyy/MM/dd 05:00:00"));
            DtStart = dt.AddDays(-1);
            DtEnd = dt;

            Stop = false;
            CountValue = -1;

            SetCommand = "";
        }

        public MySQLite(string Line, string DataBaseName = "Txn_Data.db") : this()
        {
            DB_Line = Line;
            DB_Name = DataBaseName;
        }

        public void SetDateTime(DateTime Start, DateTime End)
        {
            DtStart = Start;
            DtEnd = End;
        }

        public void ReLoad()
        {
            DB_Line_Title = ReadDt.ToString("yyyy.MM") + "月";
            DB_FullName = $"{DB_Line_Title}_{DB_Line}_{DB_Name}";
            DB_FullPath = $@"{DB_Path}{DB_Line}\";
            DB_Connect = DB_Title + DB_FullPath + DB_FullName;
            MyCommond.CheckFolder(DB_FullPath);
        }

        public Boolean Check_DB()
        {
            if (MyCommond.CheckFile(DB_FullPath + DB_FullName))
            {
                return true; 
            }
            else
            {
                MyCommond.WriteLog(ThisReceive, "未發現資料庫");
                return false; 
            }
            
        }

        public Boolean CreateDataBaseFile(string[] CommandArray)
        {
            if (SetCommand == "" && CommandArray == null) return false;
            if (File.Exists(DB_FullPath + DB_FullName)) return false;
            Status = StatusEnum.Creating;
            using (var connection = new SQLiteConnection(DB_Connect))
            {
                connection.Open();

                var command = connection.CreateCommand();

                if (SetCommand != "")
                {
                    command.CommandText = SetCommand;
                    command.ExecuteNonQuery();
                }
                else if (CommandArray != null)
                {
                    for (int i = 0; i < CommandArray.Length; i++)
                    {
                        command.CommandText = CommandArray[i];
                        command.ExecuteNonQuery();
                    }
                }
                connection.Close();
            }
            SetCommand = "";
            Status = StatusEnum.Finish;
            return true;
        }

        public Boolean Insert_(string ColumnTitle, string[] Data, int JumpTitle = 0)
        {
            Status = StatusEnum.Inserting;
            if (SetCommand == "") return false;
            if (!Check_DB()) return false;

            try
            {

                using (var connection = new SQLiteConnection(DB_Connect))
                {
                    connection.Open();
                    var command = connection.CreateCommand();

                    MaxCount = Data.Count();
                    MinCount = 0;
                    CountValue = 0;
                    string[] TitleSplit = ColumnTitle.Split(',');
                    MyCommond.WriteLog(ThisReceive, $"開始匯入營運日:{DtStart.ToString("yyyy/MM/dd")}");
                    foreach (var item in Data)
                    {
                        if (CountValue++ < JumpTitle) continue;
                        var jtem = item.Split(',');
                        var thisTable = CheckTableResult(jtem[0]);
                        command.CommandText = string.Format(SetCommand, thisTable);

                        for (int i = 0; i < TitleSplit.Count(); i++)
                        {
                            command.Parameters.AddWithValue($"${TitleSplit[i]}", jtem[i]);
                        }
                        command.ExecuteNonQuery();
                    }
                    connection.Close();
                }
                
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"Error! {ex.Message}");
            }
            Status = StatusEnum.Finish;
            CountValue = -1;
            return true;
        }

        public Boolean Insert(string ColumnTitle, string[] Data, int JumpTitle = 0)
        {
            Status = StatusEnum.Inserting;
            if (SetCommand == "") return false;
            if (!Check_DB()) return false;
            using (var connection = new SQLiteConnection(DB_Connect))
            {
                connection.Open();
                var command = connection.CreateCommand();

                command.CommandText = SetCommand;

                MaxCount = Data.Count();
                MinCount = 0;
                CountValue = 0;
                string[] TitleSplit = ColumnTitle.Split(',');
                foreach (var item in Data)
                {
                    if (CountValue++ < JumpTitle) continue;
                    var jtem = item.Split(',');
                    for (int i = 0; i < TitleSplit.Count(); i++)
                    {
                        command.Parameters.AddWithValue($"${TitleSplit[i]}", jtem[i]);
                    }
                    command.ExecuteNonQuery();
                }
                connection.Close();
            }
            Status = StatusEnum.Finish;
            CountValue = -1;
            return true;
        }

        // 先停用
        private Boolean Update(int id, string userName)
        {
            Status = StatusEnum.Updating;
            if (SetCommand == "") return false;
            using (var connection = new SQLiteConnection(DB_Connect))
            {
                connection.Open();
                var command = connection.CreateCommand();

                command.CommandText = SetCommand;
                
                command.Parameters.AddWithValue("$userName", userName);
                command.Parameters.AddWithValue("id", id);
                command.ExecuteNonQuery();
            }
            Status = StatusEnum.Finish;
            return true;
        }

        // 先停用
        private Boolean Delete(int id)
        {
            Status = StatusEnum.Deleting;
            if (SetCommand == "") return false;
            using (var connection = new SQLiteConnection(DB_Connect))
            {
                connection.Open();
                var command = connection.CreateCommand();

                command.CommandText = SetCommand;
                
                command.Parameters.AddWithValue("id", id);
                command.ExecuteNonQuery();
            }
            Status = StatusEnum.Finish;
            return true;
        }

        public List<string[]> Read()
        {
            Status = StatusEnum.Reading;
            if (SetCommand == "") return null;
            if (!Check_DB()) return null;
            List<string[]> vs = new List<string[]>();
            using (var connection = new SQLiteConnection(DB_Connect))
            {
                connection.Open();

                var command = connection.CreateCommand();

                command.CommandText = SetCommand;
                
                using (var reader = command.ExecuteReader())
                {
                    CountValue = 0;
                    while (reader.Read())
                    {
                        if (Stop)
                        {
                            CountValue = -1;
                            Stop = false;
                            break;
                        }
                        if (CountValue >= 3000000)
                        {
                            MyCommond.WriteLog(ThisReceive, "資料超過3千萬筆，強制停止");
                            CountValue = -1;
                            Stop = false;
                            break;
                        }

                        var ColumnCount = reader.FieldCount;
                        if (!(ColumnCount > 0)) continue;

                        CountValue++;
                        string[] toS = new string[ColumnCount];
                        for (int i = 0; i < ColumnCount; i++) 
                        { 
                            switch (reader.GetFieldType(i).Name)
                            {
                                case "String": /*---*/ toS[i] += Convert.ToString(reader.GetString(i));     break;
                                case "DateTime": /*-*/ toS[i] += Convert.ToString(reader.GetDateTime(i));   break;
                                case "Int16": /*----*/ toS[i] += Convert.ToString(reader.GetInt16(i));      break;
                                case "Int32": /*----*/ toS[i] += Convert.ToString(reader.GetInt32(i));      break;
                                case "Int64": /*----*/ toS[i] += Convert.ToString(reader.GetInt64(i));      break;
                                case "Byte": /*-----*/ toS[i] += Convert.ToString(reader.GetByte(i));       break;
                                case "Char": /*-----*/ toS[i] += Convert.ToString(reader.GetChar(i));       break;
                                case "Boolean": /*--*/ toS[i] += Convert.ToString(reader.GetBoolean(i));    break;
                                case "Double": /*---*/ toS[i] += Convert.ToString(reader.GetDouble(i));     break;
                                case "Float": /*----*/ toS[i] += Convert.ToString(reader.GetFloat(i));      break;
                                case "Guid": /*-----*/ toS[i] += Convert.ToString(reader.GetGuid(i));       break;
                                case "Object": /*---*/ toS[i] += Convert.ToString(reader.GetValue(i));      break;
                                default: MyCommond.WriteLog(ThisReceive, reader.GetFieldType(i).Name); break;
                            }
                        }
                        vs.Add(toS);
                    }
                }
                connection.Close();
            }
            CountValue = -1;
            Stop = false;
            Status = StatusEnum.Finish;
            return vs;
        }

        private string CheckTableResult(string TxnType)
        {
            if(DB_Line == "Mrt")
            {
                switch (Convert.ToInt16(TxnType))
                {
                    case 20:
                    case 151:
                    case 153: return $"Entry"; // Entry
                    case 1:
                    case 4:
                    case 8:
                    case 13:
                    case 152:
                    case 154: return $"Exit"; // Exit
                    case 2: return $"AddValue"; // AddValue
                    case 11: return $"AddValueCancel"; // AddValueCancel
                    case 7: return $"SaleCard"; // SaleCard
                    case 9: return $"SaleCardCancel"; // SaleCardCancel
                    case 190:
                    case 193: return $"SaleTicket"; // SaleTicket
                    case 191:
                    case 194: return $"SaleTicketRefound"; // SaleTicketRefound
                    default: return $"Other"; // Other
                }
            }

            return "";
        }

        public SQLiteConnection DB_connection { get; set; }
        public SQLiteCommand DB_command { get; set; }

        public void Locol_Open()
        {
            DB_connection = new SQLiteConnection(DB_Connect);
            DB_connection.Open();
            DB_command = DB_connection.CreateCommand();
        }

        public void Locol_Close()
        {
            DB_connection.Close();
        } 

        public SQLiteDataReader Locol_Read()
        {
            Status = StatusEnum.Reading;
            if (SetCommand == "") return null;
            if (!Check_DB()) return null;

            using (var connection = new SQLiteConnection(DB_Connect))
            {
                connection.Open();

                var command = connection.CreateCommand();

                command.CommandText = SetCommand;

                using (var reader = command.ExecuteReader())
                {
                    CountValue = 0;
                    if (reader.Read())
                    {
                        CountValue = -1;
                        Stop = false;
                        Status = StatusEnum.Finish;
                        return reader;
                    }
                }
                connection.Close();
            }

            CountValue = -1;
            Stop = false;
            Status = StatusEnum.Finish;
            return null;
        }

        public void test_1()
        {
            SetCommand = @"select * from Exit where TXN_TIMESTAMP > '2023-08-16 06:00:00' and TXN_TIMESTAMP < '2023-08-16 06:10:00'";
            var re = Read();

            List<TxnTransaction> txns = new List<TxnTransaction>();

            foreach (var item in re)
            {

            }
        }

        private string Var2String(SQLiteDataReader reader, int column)
        {
            var a1 = reader.GetFieldType(column).Name;
            string ReString = "";
            try
            {
                switch (reader.GetFieldType(column).Name)
                {
                    case "String": /*---*/ ReString += $"{Convert.ToString(reader.GetString(column))},"; break;
                    case "DateTime": /*-*/ ReString += $"{Convert.ToString(reader.GetDateTime(column))},"; break;
                    case "Int16": /*----*/ ReString += $"{Convert.ToString(reader.GetInt16(column))},"; break;
                    case "Int32": /*----*/ ReString += $"{Convert.ToString(reader.GetInt32(column))},"; break;
                    case "Int64": /*----*/ ReString += $"{Convert.ToString(reader.GetInt64(column))},"; break;
                    case "Byte": /*-----*/ ReString += $"{Convert.ToString(reader.GetByte(column))},"; break;
                    case "Char": /*-----*/ ReString += $"{Convert.ToString(reader.GetChar(column))},"; break;
                    case "Boolean": /*--*/ ReString += $"{Convert.ToString(reader.GetBoolean(column))},"; break;
                    case "Double": /*---*/ ReString += $"{Convert.ToString(reader.GetDouble(column))},"; break;
                    case "Float": /*----*/ ReString += $"{Convert.ToString(reader.GetFloat(column))},"; break;
                    case "Guid": /*-----*/ ReString += $"{Convert.ToString(reader.GetGuid(column))},"; break;
                    case "Object": /*---*/ ReString += $"{Convert.ToString(reader.GetValue(column))},"; break;
                    default: MyCommond.WriteLog(ThisReceive, $"{reader.GetFieldType(column).Name}"); break;
                }
            }
            catch (Exception ex)
            {

            }
            return ReString;
        }

    }

    internal class TableNum
    {
        public int Id { get; set; }
        public string Line { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }

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
                $"{LAST_UTILISATION_DATE}";

        public string ExportTitle() => $"CARD_TXN_TYPE_ID," +
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
            $"LAST_UTILISATION_DATE";

        #endregion
    }

}
