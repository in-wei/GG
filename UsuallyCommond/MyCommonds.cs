using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//--
using System.IO;
using System.Windows.Forms;
using UsuallyCommond.MyEnum;
using System.Diagnostics;
using System.Management;
using System.Collections;
using System.Runtime.InteropServices;
using ClosedXML.Excel;
using System.Reflection;
using System.ComponentModel;
using Aspose.Zip.SevenZip;
using System.Net;
//using System.IO.Compression;
//using Aspose.Cells;

public class ExportConfig
{
    public string Path1 { get; set; }
    public DateTime DTime { get; set; }
    public YearType yearType { get; set; }
    public string FirstName { get; set; }
    public string FileName { get; set; }
    public string SubName { get; set; }
    public string OpLine { get; set; }
    public string LastName { get; set; }

    public ExportConfig()
    {
        Path1 = "";
        DTime = new DateTime();
        yearType = YearType.AD;
        FirstName = "";
        FileName = "";
        SubName = "";
        OpLine = "";
        LastName = "csv";
    }

    /// <summary>
    /// 只有檔名，為了可以重新命名，單獨使用需自行加入副檔名。
    /// </summary>
    /// <returns></returns>
    public override string ToString()
    {
        string Respon = "";

        string filePath = Path1;
        string fileName1 = ""; // 時間
        string fileName2 = FirstName;
        string fileName3 = FileName;
        string fileName4 = SubName;
        string fileName5 = OpLine;

        if (DTime != new DateTime())
        {
            if (yearType == YearType.AD)
            {
                fileName1 = DTime.ToString("yyyyMMdd");
            }
            else if (yearType == YearType.RC)
            {
                var mYear = ((Convert.ToInt32(DTime.ToString("yyyy")) > 1911) ? (Convert.ToInt32(DTime.ToString("yyyy")) - 1911) : Convert.ToInt32(DTime.ToString("yyyy"))).ToString();
                var mMonth = DTime.ToString("MM");
                var mDay = DTime.ToString("dd");
                fileName1 = $"{mYear}{mMonth}{mDay}";
            }
            else
            {

            }
        }

        if (fileName1.Length > 0) { if (Respon.Length > 0) { Respon += "_"; } /*-*/ Respon += fileName1; }
        if (fileName2.Length > 0) { if (Respon.Length > 0) { Respon += "_"; } /*-*/ Respon += fileName2; }
        if (fileName3.Length > 0) { if (Respon.Length > 0) { Respon += "_"; } /*-*/ Respon += fileName3; }
        if (fileName4.Length > 0) { if (Respon.Length > 0) { Respon += "_"; } /*-*/ Respon += fileName4; }
        if (fileName5.Length > 0) { if (Respon.Length > 0) { Respon += "_"; } /*-*/ Respon += fileName5; }

        return Respon;
    }
}

public class FileLastName
{
    public bool csv { get; set; }
    public bool xlsx { get; set; }
    public bool xls { get; set; }
    public bool txt { get; set; }
    public bool xlsm { get; set; }
    public bool xml { get; set; }

    public FileLastName()
    {
        csv = false;
        xlsx = false;
        xls = false;
        txt = false;
        xlsm = false;
        xml = false;
    }
}

public class FolderFile
{
    public int No { get; set; }
    public string PathOrFilename { get; set; }

    public FolderFile(int _no = 0, string _path = "")
    {
        No = _no;
        PathOrFilename = _path;
    }

    public override string ToString()
    {
        return PathOrFilename;
    }

    public string ExportTitle_zhTW()
    {
        return "號,路徑";
    }

}

public class MyCommonds
{
    #region 路徑
    public string Path_Program = System.IO.Directory.GetCurrentDirectory() + @"\";
    public string Path_Analyze = @"Analyze\";
    public string Path_Config = @"Config\";
    public string Path_Report = @"Report\";
    public string Path_Template = @"Template\";
    public string Path_Tidy = @"Tidy\";
    public string Path_Verify = @"Verify\";
    public string Path_FileZilla_root = @"C:\ftproot\";
    public string Path_FileZilla_History = @"history\";
    public string Path_FileZilla_TclTxn = @"TCLTxn\";

    public string Path_Date = @"{0}\";
    #endregion

    public Language ExportLanguage = Language.English;
    //private string ThisReceive;

    private int Time_1s = 1000;

    public MyCommonds()
    {
        //ThisReceive = "MyCommond";
        CheckFolder(Path_Program + Path_Analyze + @"Old\");
        CheckFolder(Path_Program + Path_Config);
        CheckFolder(Path_Program + Path_Report + @"Old\");
        CheckFolder(Path_Program + Path_Template + @"Old\");
        //CreatFolder(Path_Program + Path_Tidy + @"Old\");
        //CreatFolder(Path_Program + Path_Verify + @"Old\");
    }

    public MyCommonds(bool CreateAnalyze = false, bool CreateConfig = false, bool CreateReport = true, bool CreateTemplate = false)
    {
        //ThisReceive = "MyCommond";
        if (CreateAnalyze) CheckFolder(Path_Program + Path_Analyze + @"Old\");
        if (CreateConfig) CheckFolder(Path_Program + Path_Config);
        if (CreateReport) CheckFolder(Path_Program + Path_Report + @"Old\");
        //if (CreateTemplate) CreatFolder(Path_Program + Path_Template + @"Old\");
        //CreatFolder(Path_Program + Path_Tidy + @"Old\");
        //CreatFolder(Path_Program + Path_Verify + @"Old\");
    }

    #region 委派

    /// <summary>
    /// 委派動作到Form上的物件。
    /// this.button1.InvokeIfRequired(() => { this.button1.Enabled = TrueOrFalse; });
    /// </summary>
    /// <param name="control">要委派的物件。</param>
    /// <param name="action">要這個物件執行什麼動作。</param>
    public void InvokeIfRequired(Control control, MethodInvoker action)
    {
        if (control.InvokeRequired) { control.Invoke(action); }
        else { action(); }
    }

    /// <summary>
    /// 選擇檔案。
    /// </summary>
    /// <returns></returns>
    public string SelectFileForm(FileLastName fileLastName)
    {
        OpenFileDialog openTxt = new OpenFileDialog();//實例化打開對話框對像
        //openTxt.Filter = "CSV Files|*.csv|Excel File|*.xlsx|Excel File|*.lsx|All Files|*.*";

        string FilterString = "";

        if (fileLastName.txt) /*--*/ { FilterString += @"文字文件|*.txt|"; }
        if (fileLastName.csv) /*--*/ { FilterString += @"CSV (逗號分隔)|*.csv|"; }
        if (fileLastName.xlsx) /*-*/ { FilterString += @"Excel 活頁簿|*.xlsx|"; }
        if (fileLastName.xls) /*--*/ { FilterString += @"Excel 97-2003 活頁簿|*.xls|"; }
        if (fileLastName.xlsm) /*-*/ { FilterString += @"Excel 啟用聚集的活頁簿|*.xlsm|"; }
        if (fileLastName.xml) /*--*/ { FilterString += @"XML 試算表2003|*.xml|"; }

        FilterString += @"All Files|*.*";
        openTxt.Filter = FilterString;

        openTxt.Multiselect = false;//設定打開對話框中不能多選
        if (openTxt.ShowDialog() == DialogResult.OK)//判斷是否選擇了文件
        {
            return openTxt.FileName;//顯示選擇的文字文件
        }
        return "";
    }

    /// <summary>
    /// 選擇資料夾。
    /// </summary>
    /// <returns></returns>
    public string SelectFolderForm()
    {
        System.Windows.Forms.FolderBrowserDialog ff = new System.Windows.Forms.FolderBrowserDialog();
        
        if (ff.ShowDialog() == DialogResult.OK)
        {
            return ff.SelectedPath;
        }

        return "";
    }

    #endregion

    #region 檢查檔案、文件、資料

    /// <summary>
    /// 讓檢察文件更直覺。
    /// </summary>
    /// <param name="FullFilePath"></param>
    /// <returns></returns>
    public bool CheckFile(string FullFilePath) => File.Exists(FullFilePath);

    /// <summary>
    /// 檢查該路徑是否可以抵達，無法抵達的話將會創建完整的路徑。
    /// </summary>
    /// <param name="Path">要檢查的路徑。</param>
    public bool CheckFolder(string Path) { if (!Directory.Exists(Path)) { Directory.CreateDirectory(Path); } return true; }

    /// <summary>
    /// 檢測是否已有該檔案名稱在指定路徑上，假如有就重新將檔案命名。
    /// </summary>
    /// <param name="Path">檔案要放的路徑。</param>
    /// <param name="FileName">該檔案的名稱。</param>
    /// <param name="LastName">該檔案的副檔名。</param>
    /// <returns>完整的檔案名稱。Txn.csv</returns>
    public string CheckAndReName(string Path, string FileName, string LastName)
    {
        string Result = $"{FileName}.{LastName}";
        int RoopCount = 0;
        CheckFolder(Path);
        if (CheckFile(Path + $"{FileName}.{LastName}"))
        {
            while (CheckFile(Path + $"{FileName}_{RoopCount}.{LastName}"))
            {
                RoopCount++;
            }
            Result = $"{FileName}_{RoopCount}.{LastName}";
        }
        return Result;
    }

    /// <summary>
    /// 將檔案搬至其他地方存放。(如果有重複的檔名，會自動重新命名)
    /// </summary>
    /// <param name="OldPath">原檔案存放位置。</param>
    /// <param name="NewPath">新存放位置。</param>
    /// <param name="FileName">原檔案名稱。</param>
    /// <returns></returns>
    /// <exception cref="Exception"></exception>
    public bool MoveFile(string OldPath, string NewPath, string FileName)
    {
        if (!CheckFolder(OldPath)) { throw new Exception("資料夾創建失敗!"); }
        if (!CheckFile($@"{OldPath}\{FileName}")) { throw new Exception($@"未發現需要被搬移的檔案。\n路徑:{OldPath}\{FileName}"); }
        var FileName_Split_1 = FileName.Split('.');
        var FileName_LastName = FileName_Split_1[FileName_Split_1.GetLength(0) - 1];
        var FileName1 = "";
        for (int i = 0; i < FileName_Split_1.GetLength(0) - 1; i++) { FileName1 += FileName_Split_1[i] + ((i == FileName_Split_1.GetLength(0) - 2) ? "" : "."); }
        string tempFileName = $"{FileName1}.{FileName_LastName}";
        int RoopCount = 0;
        CheckFolder(NewPath);
        while (CheckFile(NewPath + tempFileName)) { tempFileName = $"{FileName1}_{RoopCount++}.{FileName_LastName}"; }

        File.Move($@"{OldPath}\{FileName}", $@"{NewPath}\{tempFileName}");

        if (CheckFile($@"{NewPath}\{tempFileName}")) { return true; }

        return false;
    }

    public bool B_CheckLine(string CodeString)
    {
        Boolean tf = false;
        string msgStr = "比對資料： {0,6} ,被比對資料1： {1,3} ,被比對資料2： {2,3} ,被比對資料3：{3}\t ,是否相符：{4}";
        //tf = OperationLineList.Find(x => x.CodeName == CodeString || x.Chinese == CodeString) != null;
        return tf;
    }

    public string S_CheckLine(string CodeString)
    {
        string msgStr = "比對資料： {0,6} ,被比對資料1： {1,3} ,被比對資料2： {2,3} ,被比對資料3：{3}\t ,是否相符：{4}";
        string tf = "'";

        return tf;
    }

    public bool B_CheckStation(string Station)
    {
        Boolean tf = false;
        string msgStr = "比對資料： {0,6} ,被比對資料1： {1,3} ,被比對資料2： {2,3} ,被比對資料3：{3}\t ,是否相符：{4}";
        return tf;
    }

    public int GetStationRow(Control GetThis, string Station, bool IsEntey)
    {
        Int16 ReturnCol = 0;

        if (B_CheckStation(Station))
        {
            byte mOffset;
            if (IsEntey) mOffset = 0;
            else mOffset = 1;

        }
        else { ReturnCol = -1; }

        return ReturnCol;
    }

    public bool Num2Boolean(int Num)
    {
        return Num == 1;
    }

    public ExecutionMode Config_Using_Mode(string s)
    {
        if (s == "Normal") return ExecutionMode.Normal;
        else if (s == "Debug") return ExecutionMode.Debug;
        else return ExecutionMode.Simple;
    }

    public Language Config_ExportLanguage(string s)
    {
        if (s == "Number") return Language.Number;
        else if (s == "Num") return Language.Number;
        else if (s == "Taiwanese") return Language.zhTW;
        else if (s == "zhTW") return Language.zhTW;
        else if (s == "TW") return Language.zhTW;
        else if (s == "Tw") return Language.zhTW;
        else return Language.English;
    }

    /// <summary>
    /// 完整掃瞄指定的目錄。
    /// </summary>
    /// <param name="fPath"></param>
    /// <param name="fs"></param>
    /// <returns></returns>
    public string[] ScanFolder(FolderFile fPath, ref List<FolderFile> fs)
    {
        string thisPath = fPath.PathOrFilename + @"\";
        var ff = System.IO.Directory.GetFileSystemEntries(thisPath, @"*.*");
        int i = 0;
        foreach (string fname in ff)
        {
            FolderFile f1 = new FolderFile(++i, fname);
            fs.Add(f1);
            try
            {
                ScanFolder(f1, ref fs);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Not Folder");
            }
        }

        return null;
    }

    /// <summary>
    /// 完整掃瞄指定的目錄中的檔案數量，並將掃描的檔案名稱輸出一份報告。
    /// </summary>
    /// <param name="Path">要檢查的路徑。</param>
    /// <param name="MinCount">最少需要有幾個檔案。</param>
    /// <param name="Export">是否需要輸出檔案。(預設:否)</param>
    /// <returns>輸出是否。</returns>
    public bool ScanFolderAndExport(string Path, int MinCount = 0, bool Export = false)
    {
        List<FolderFile> folderFiles = new List<FolderFile>();
        ScanFolder(new FolderFile(0, Path), ref folderFiles);

        string _Path = $@"{Path_Program}{Path_Report}";
        CheckFolder(_Path);
        string _FileName = $@"{DateTime.Now.ToString("yyyyMMdd")}_scanPath";
        string _LastName = "txt";
        string FileName = CheckAndReName(_Path, _FileName, _LastName);
        string _FullPath = _Path + FileName;

        if (Export)
        {
            for (int i = 0; i < folderFiles.Count; i++)
            {
                ExportData(_FullPath, Encoding.UTF8, folderFiles[i].ToString());
            }
        } // 是否輸出檢測到的檔案
        
        if (folderFiles.Count > MinCount) return true;

        return false;
    }

    public bool CheckSimilarFile(ref string fullPathName)
    {
        try
        {
            string[] p1 = fullPathName.Split('\\');
            string path = "";
            for (int i = 0; i < p1.GetLength(0) - 1; i++)
            {
                path += p1[i] + @"\";
            }
            string[] n0 = p1[p1.GetLength(0) - 1].Split('.');
            string n1 = n0[0];
            string n1_1 = n1;
            string n2 = n0[1];
            int subInt = 0;
            while (CheckFile(path + $"{n1_1}.{n2}"))
            {
                n1_1 = $"{n1}_{subInt++}";
            }
            if (--subInt == 0) { return true; }
            fullPathName = path + $"{n1}_{--subInt}.{n2}";
            return true;
        }
        catch (Exception ex)
        {
            return false;
        }
    }

    #endregion

    #region 資料整理轉換

    /// <summary>
    /// 將符合欄位長度的資料寫入。並依照欄位的型態寫入資料。
    /// </summary>
    /// <typeparam name="TResult">賦予的型態。</typeparam>
    /// <param name="Receive">使用他的目錄。</param>
    /// <param name="Path">檔案的完整路徑及檔名。</param>
    /// <param name="DataEow">一開始需要跳過多少行。預設為0。</param>
    /// <returns>載入完成的變數。</returns>
    public List<TResult> ToViewModel<TResult>(string Receive, string Path, string TResultTile, int DataEow = 0, bool GetLog = false) where TResult : class, new()
    {
        List<TResult> Result_List = new List<TResult>();
        string[] Data;
        WriteLog(Receive, $"載入資料的路徑: {Path}");
        if (!CheckFile(Path)) return Result_List;
        WriteLog(Receive, $"確認有此檔案，開始載入");
        Data = File.ReadAllLines(Path, Encoding.Default);

        Type type = typeof(TResult);

        WriteLog(Receive, $"資料總比數:{Data.Length}");

        int RunCount = 0;
        string DataTile = "";
        foreach (var item in Data)
        {
            //if (GetLog) { WriteLog(ThisReceive, $"LoadLine:{item.ToString()}"); }
            if (DataEow > RunCount)
            {
                DataTile = item;
                if (GetLog) WriteLog(Receive, $"此為資料欄名稱..LoadLine:{item.ToString()}");
                goto ContinueNext;
            }
            if (GetLog) WriteLog(Receive, $"此為資料內容..LoadLine:{item.ToString()}");
            if (item == "") {  goto ContinueNext; }
            try
            {
                string[] itemSplit = item.Split(',');
                Result_List.Add(T2<TResult>(Receive, TResultTile, DataTile, itemSplit));
            }
            catch (Exception ex)
            {
                WriteLog(Receive, $"ERROR! \n{ExportErrorMessageToLog(ex).Message}");
            }

        ContinueNext:
            RunCount++;
        }
        //GetThis.WriteLog(ExecutionMode.Simple, "載入結束");
        return Result_List;
    }

    public TResult T2<TResult>(string Receive, string TResultTile, string DataTitle, string[] RowData) where TResult : class, new()
    {
        TResult tr = new TResult();
        Type type = typeof(TResult);
        int itemCount = 0;
        var mTitle = TResultTile.Split(',');
        var dTitle = DataTitle.Split(',');
        foreach (var item in RowData)
        {
            bool ReLoop = false;
            int LoopCount = 0;
            StartInstall:
            try
            {
                var tt = dTitle[itemCount];
                if (tt == null) continue;
                var TitleCount = 0;
                foreach (var jtem in mTitle)
                {
                    if (jtem == tt) break;
                    TitleCount++;
                }
                if (mTitle.Length == TitleCount)
                {
                    itemCount++;
                    break;
                }
                if (item != "")
                {
                    string ggg = type.GetProperties()[TitleCount].PropertyType.Name;
                         if (ggg == "DateTime") type.GetProperties()[TitleCount].SetValue(tr, Convert.ToDateTime(item), null);
                    else if (ggg == "Boolean" ) type.GetProperties()[TitleCount].SetValue(tr, Convert.ToBoolean(item) , null);
                    else if (ggg == "String"  ) type.GetProperties()[TitleCount].SetValue(tr, Convert.ToString(item)  , null);
                    else if (ggg == "Char"    ) type.GetProperties()[TitleCount].SetValue(tr, Convert.ToChar(item)    , null);
                    else if (ggg == "Int16"   ) type.GetProperties()[TitleCount].SetValue(tr, Convert.ToInt16(item)   , null);
                    else if (ggg == "Int32"   ) type.GetProperties()[TitleCount].SetValue(tr, Convert.ToInt32(item)   , null);
                    else if (ggg == "Int64"   ) type.GetProperties()[TitleCount].SetValue(tr, Convert.ToInt64(item)   , null);
                }
                itemCount++;
                ReLoop = false;
            }
            catch (Exception ex)
            {
                WriteLog(Receive, $"ERROR! \n{ExportErrorMessageToLog(ex).Message}");
                ReLoop = true;
            }
            finally
            {

            }
            if (ReLoop)
            {
                LoopCount++;
                goto StartInstall;
            }
        }
        return tr;
    }

    public string ExportViewTitle<TResult>()
    {
        string ss = "";
        //使用 reflection 將物件屬性取出當作工作表欄位名稱
        foreach (var item in typeof(TResult).GetProperties())
        {
            //#region - 可以使用 DescriptionAttribute 設定，找不到 DescriptionAttribute 時改用屬性名稱 -
            ////可以使用 DescriptionAttribute 設定，找不到 DescriptionAttribute 時改用屬性名稱
            if (item.GetCustomAttributes(typeof(DescriptionAttribute)) is DescriptionAttribute description)
            {
                ss += $"{description.Description},";
                continue;
            }
            ss += $"{item.Name},";
            //#endregion
        }
        return "";
    }

    public int StringCount(string str)
    {
        int english = 0, chinese = 0;

        foreach (var item in str)
        {
            //Console.WriteLine($"文字:{item}, {(byte)item}");
            if (item >= 0x3000 && item <= 0x9FFF)
            {
                chinese++;
            }
            else
            {
                english++;
            }
        }
        
        return english + (chinese * 2);
    }
    /** public string[,] TResultToArray<TResult>(string Receive, List<TResult> InputData)
    {
        WriteLog(Receive, $"將資料表轉換成多行陣列 - 開始");

        var Count_R = InputData.Count;
        var Count_C = InputData[0].ToString().Split(',').Length;
        string[,] OutputData = new string[Count_R, Count_C];
        int Run_R = 0;
        foreach (var item in InputData)
        {
            var LineData = item.ToString().Split(',');
            for (int Run_C = 0; Run_C < LineData.Length; Run_C++)
            {
                OutputData[Run_R, Run_C] = Convert.ToString(LineData[Run_C]);
            }
        }

        WriteLog(Receive, $"將資料表轉換成多行陣列 - 結束");
        return OutputData;
    }
    */
    /** public int[,] TResultToArray<TResult>(string Receive, List<TResult> InputData)
    {
        WriteLog(Receive, $"將資料表轉換成多行陣列 - 開始");

        var Count_R = InputData.Count;
        var Count_C = InputData[0].ToString().Split(',').Length;
        int[,] OutputData = new int[Count_R, Count_C];
        int Run_R = 0;
        foreach (var item in InputData)
        {
            var LineData = item.ToString().Split(',');
            for (int Run_C = 0; Run_C < LineData.Length; Run_C++)
            {
                OutputData[Run_R, Run_C] = Convert.ToInt32(LineData[Run_C]);
            }
        }

        WriteLog(Receive, $"將資料表轉換成多行陣列 - 結束");
        return OutputData;
    }
    */
    public int[,] TResultToArray<TResult>(string Receive, TResult InputData)
    {
        WriteLog(Receive, $"將資料表轉換成單行陣列 - 開始");

        var Count_C = InputData.ToString().Split(',').Length;
        int[,] OutputData = new int[1, Count_C];
        var LineData = InputData.ToString().Split(',');
        for (int Run_C = 0; Run_C < LineData.Length; Run_C++)
        {
            OutputData[0, Run_C] = Convert.ToInt32(LineData[Run_C]);
        }

        WriteLog(Receive, $"將資料表轉換成單行陣列 - 結束");
        return OutputData;
    }

    public int[,] TresultToArray<Tresult>(string Receive, List<Tresult> InputData)
    {
        var DataRow = InputData.Count;
        var DataColumn = nameof(Tresult).Split(',').GetLongLength(0);
        int[,] respone = new int[DataRow, DataColumn];

        int Run_R = 0;
        int Run_C = 0;

        foreach (var item in InputData)
        {
            var s1 = item.ToString().Split(',');
            foreach (var jtem in s1)
            {
                respone[Run_R, Run_C] = Convert.ToInt32(s1[Run_C]);
                Run_C++;
            }
            Run_R++;
        }

        return respone;
    }

    #region Excel欄位轉換
    public Int32 ExcelColumnNum(string str)
    {
        Int32 ChangeNum = 0;
        int Len = str.Length - 1;
        foreach (var item in str)
        {
            int temp = EnglishToNumber(item.ToString());
            if (Len > 0)
            {
                ChangeNum += temp * Len * 26;
                Len--;
            }
            else
            {
                ChangeNum += temp;
            }
        }
        return ChangeNum;
    }

    public string ExcelColumnStr(Int32 num)
    {
        string ChangeNum = "";
        if (num > 26)
        {
            int String_0 = num % 26;
            int String_1 = num / 26;
            int String_2 = 0;
            if (String_1 > 26)
            {
                String_2 = String_1 / 26;
                String_1 = String_1 % 26;
            }
            ChangeNum = NumberToEnglish(String_2) + NumberToEnglish(String_1) + NumberToEnglish(String_0);
        }
        else
        {
            ChangeNum = NumberToEnglish(num);
        }

        return ChangeNum;
    }

    private int EnglishToNumber(string str)
    {
        Int32 ChangeNum = 0;
        switch (str)
        {
            case "A": ChangeNum = 1; break;
            case "B": ChangeNum = 2; break;
            case "C": ChangeNum = 3; break;
            case "D": ChangeNum = 4; break;
            case "E": ChangeNum = 5; break;
            case "F": ChangeNum = 6; break;
            case "G": ChangeNum = 7; break;
            case "H": ChangeNum = 8; break;
            case "I": ChangeNum = 9; break;
            case "J": ChangeNum = 10; break;
            case "K": ChangeNum = 11; break;
            case "L": ChangeNum = 12; break;
            case "M": ChangeNum = 13; break;
            case "N": ChangeNum = 14; break;
            case "O": ChangeNum = 15; break;
            case "P": ChangeNum = 16; break;
            case "Q": ChangeNum = 17; break;
            case "R": ChangeNum = 18; break;
            case "S": ChangeNum = 19; break;
            case "T": ChangeNum = 20; break;
            case "U": ChangeNum = 21; break;
            case "V": ChangeNum = 22; break;
            case "W": ChangeNum = 23; break;
            case "X": ChangeNum = 24; break;
            case "Y": ChangeNum = 25; break;
            case "Z": ChangeNum = 26; break;
            default: ChangeNum = 0; break;
        }

        return ChangeNum;
    }

    private string NumberToEnglish(int str)
    {
        string ChangeNum = "";
        switch (str)
        {
            case 1: ChangeNum = "A"; break;
            case 2: ChangeNum = "B"; break;
            case 3: ChangeNum = "C"; break;
            case 4: ChangeNum = "D"; break;
            case 5: ChangeNum = "E"; break;
            case 6: ChangeNum = "F"; break;
            case 7: ChangeNum = "G"; break;
            case 8: ChangeNum = "H"; break;
            case 9: ChangeNum = "I"; break;
            case 10: ChangeNum = "J"; break;
            case 11: ChangeNum = "K"; break;
            case 12: ChangeNum = "L"; break;
            case 13: ChangeNum = "M"; break;
            case 14: ChangeNum = "N"; break;
            case 15: ChangeNum = "O"; break;
            case 16: ChangeNum = "P"; break;
            case 17: ChangeNum = "Q"; break;
            case 18: ChangeNum = "R"; break;
            case 19: ChangeNum = "S"; break;
            case 20: ChangeNum = "T"; break;
            case 21: ChangeNum = "U"; break;
            case 22: ChangeNum = "V"; break;
            case 23: ChangeNum = "W"; break;
            case 24: ChangeNum = "X"; break;
            case 25: ChangeNum = "Y"; break;
            case 26: ChangeNum = "Z"; break;
            default: ChangeNum = ""; break;
        }

        return ChangeNum;
    }

    #endregion

    #endregion

    #region 壓縮檔處理

    /// <summary>
    /// 解壓縮檔案至某個資料夾。(會判斷有沒有檔案)(如果有重複的檔案會死去)
    /// </summary>
    /// <param name="ZipFullPath">壓縮檔完整路徑。</param>
    /// <param name="ExtractPath">解壓縮後存放的位置。</param>
    /// <param name="filePassword">解壓縮密碼。</param>
    /// <param name="overwrite">覆蓋重複檔案。</param>
    /// <returns></returns>
    /// <exception cref="Exception"></exception>
    public bool Cmd_Decompression(string Receive, string ZipFullPath, string ExtractPath = "", string filePassword = "", bool overwrite = true)
    {
        if (ZipFullPath == "") throw new Exception("未取得壓縮檔路徑。");
        if (!CheckFile(ZipFullPath)) throw new Exception("未發現檔案。");
        if (ExtractPath == "") ExtractPath = Path_Program + Path_Analyze;

        try
        {
            string CommandString = $@"""C:\Program Files\7-Zip\7z.exe"" e ""{ZipFullPath}""{((overwrite) ? " -aoa" : "")} -o""{ExtractPath}""";
            WriteLog(Receive, CommandString);
            using (Process process = new Process())
            {
                process.StartInfo.FileName = "cmd.exe";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardInput = true;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                process.StartInfo.CreateNoWindow = true;
                process.Start();
                //process.StandardInput.WriteLine("cls");
                process.StandardInput.WriteLine(CommandString);
                process.StandardInput.WriteLine("exit");
                string strOutput = process.StandardOutput.ReadToEnd();
                WriteLog(Receive, $"Show Console Message:\n{strOutput}");
                process.WaitForExit();
                process.Close();
                System.Threading.Thread.Sleep(10 * Time_1s);
            }
            return true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
    }

    /// <summary>
    /// 解壓縮檔案至某個資料夾。(會判斷有沒有檔案)(如果有重複的檔案會死去)
    /// </summary>
    /// <param name="ZipFullPath">壓縮檔完整路徑</param>
    /// <param name="ExtractPath">解壓縮後存放的位置。(預設Analyze)</param>
    /// <returns></returns>
    /// <exception cref="Exception"></exception>
    public bool M_Decompression(string ZipFullPath, string ExtractPath = "")
    {
        if (ZipFullPath == "") throw new Exception("未取得壓縮檔路徑。");
        if (!CheckFile(ZipFullPath)) throw new Exception("未發現檔案。");
        if (ExtractPath == "") ExtractPath = Path_Program + Path_Analyze;
        try
        {
            // 使用 SevenZipArchive 類加載輸入 7z (7zip) 存檔。
            using (SevenZipArchive archive = new SevenZipArchive(ZipFullPath))
            {
                // 使用 ExtractToDirectory 方法將 7zip 中的所有文件提取到目錄中。
                archive.ExtractToDirectory(ExtractPath);
                System.Threading.Thread.Sleep(5 * Time_1s);
            }
            return true;
        }
        catch (Exception ex)
        {
            throw new Exception("解壓縮異常。", ex);
        }


        return false;
    }

    public bool FileZip(List<string> fileList, string targetPath, string targetFileName, string filePassword)
    {
        try
        {
            // 若不指定目的與目的檔名則取第一個 List 當作目的檔名
            if (!string.IsNullOrEmpty(fileList.FirstOrDefault()) && string.IsNullOrEmpty(targetPath))
            {
                targetPath = Path.GetDirectoryName(fileList.FirstOrDefault());
                targetFileName = Path.GetFileNameWithoutExtension(fileList.FirstOrDefault());
            }

            using (Ionic.Zip.ZipFile dotZip = new Ionic.Zip.ZipFile())
            {
                if (!string.IsNullOrEmpty(filePassword))
                    dotZip.Password = filePassword;

                foreach (var item in fileList)
                {
                    if (File.Exists(item.ToString()))
                        dotZip.AddFile(item.ToString(), "");
                }

                dotZip.Save(string.Format(@"{0}.zip", Path.Combine(targetPath, targetFileName)));
            }
            return true;
        }
        catch (Exception ex)
        {
            return false;
        }
    }  // 壓縮

    public bool fileUnZip(string filePath, string targetPath = "", string filePassword = "")
    {
        if (targetPath == "") targetPath = Path_Program + Path_Analyze;
        using (Ionic.Zip.ZipFile DotZip = Ionic.Zip.ZipFile.Read(filePath))
        {
            if (!string.IsNullOrEmpty(filePassword))
                DotZip.Password = filePassword;

            DotZip.ExtractAll(targetPath, Ionic.Zip.ExtractExistingFileAction.OverwriteSilently);  // 解壓縮路徑  
            var result = DotZip.EntryFileNames.ToList();
        }

        return true;
    }  // 解壓縮

    public bool UnzipToDirectory(string zipFile, string outDir = "", bool overwrite = true)
    {
        if (outDir == "") outDir = Path_Program + Path_Analyze;
        try
        {
            var archive = Ionic.Zip.ZipFile.Read(zipFile);
            foreach (var file in archive.Entries)
            {
                // file.Name == "" 表示 file 為目錄
                if (file.FileName == "")
                {
                    string desPath = Path.Combine(outDir, file.FileName);
                    // 目錄不存在就要建立
                    if (!Directory.Exists(desPath))
                    {
                        Directory.CreateDirectory(desPath);
                    }
                }
                else
                {
                    // file 為檔案
                    string desFileName = Path.Combine(outDir, file.FileName);
                    // 可覆寫就直接解壓縮
                    if (overwrite)
                    {
                        file.Extract(desFileName, Ionic.Zip.ExtractExistingFileAction.OverwriteSilently);
                    }
                    else
                    {
                        // 不可覆寫就要先判斷檔案是否存在，不存在才解壓縮
                        if (!File.Exists(desFileName))
                        {
                            file.Extract(desFileName, Ionic.Zip.ExtractExistingFileAction.DoNotOverwrite);
                        }
                    }

                }
            }
            return true;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        
    }

    public static void Unzip(string ZipFullPath)
    {
        DirectoryInfo DirecInfo = new DirectoryInfo(ZipFullPath);
        if (DirecInfo.Exists)
        {
            foreach (FileInfo fileInfo in DirecInfo.GetFiles("*.7z"))
            {
                Process process = new Process();
                process.StartInfo.FileName = @"C:\Program Files\7-zip\7z.exe";
                process.StartInfo.Arguments = @" e C:\Directory\" + fileInfo.Name + @" -o C:\Directory";
                process.Start();
            }
        }
    }

    #endregion

    #region 輸出資料

    private string ErrorMessage = "";
    public Exception ExportErrorMessageToLog(Exception ex, string MethodeString = "")
    {
        lock (ErrorMessage)
        {
            System.Threading.Thread thread = new System.Threading.Thread(() =>
            {
                MessageBox.Show($"錯誤!\n{ex.Message}", MethodeString);
            });
            thread.IsBackground = true;
            thread.Start();
        }
        return ex;
    }

    public void MsgShow(string text, string title, MessageBoxButtons mbb, MessageBoxIcon mbi)
    {
        System.Threading.Thread thread = new System.Threading.Thread(() =>
        {
            MessageBox.Show(text, title, mbb, mbi);
        });
        thread.IsBackground = true;
        thread.Start();
    }

    public void ConsoleExportArray(string Receive, string Methon, int[,] result, string Title = "")
    {
        if (result == null)
        {
            WriteLog(Receive, $"{Methon,20}---Null---");
            return;
        }
        if (Title != "")
        {
            WriteLog(Receive, $"{Methon,20}---FullAdd({Title})---");
        }
        else
        {
            WriteLog(Receive, $"{Methon,20}---FullAdd---");
        }

        for (int i = 0; i < result.GetLength(0); i++)
        {
            string LineString = "";
            for (int j = 0; j < result.GetLength(1); j++)
            {
                LineString += $"{((result[i, j] != null) ? result[i, j] : 0),4}, ";
            }
            WriteLog(Receive, $"vs[{i,4}] = {LineString}");
        }
    }

    public void ConsoleExportArray(string Receive, string Methon, string[,] result, string Title = "")
    {
        if (result == null)
        {
            WriteLog(Receive, $"{Methon,20}---Null---");
            return;
        }
        if (Title != "")
        {
            WriteLog(Receive, $"{Methon,20}---FullAdd({Title})---");
        }
        else
        {
            WriteLog(Receive, $"{Methon,20}---FullAdd---");
        }

        for (int i = 0; i < result.GetLength(0); i++)
        {
            string LineString = "";
            for (int j = 0; j < result.GetLength(1); j++)
            {
                LineString += $"{((result[i, j] != null) ? result[i, j] : ""),8}, ";
            }
            WriteLog(Receive, $"vs[{i,4}] = {LineString}");
        }
    }

    public void ConsoleExportArray<Tresult>(string Receive, string Methon, List<Tresult> result, string Title = "")
    {
        if (result == null)
        {
            WriteLog(Receive, $"{Methon,20}---Null---");
            return;
        }
        if (Title != "")
        {
            WriteLog(Receive, $"{Methon,20}---FullAdd({Title})---");
        }
        else
        {
            WriteLog(Receive, $"{Methon,20}---FullAdd---");
        }

        var Data_Row = 0;
        foreach (var item in result)
        {
            WriteLog(Receive, $"vs[{Data_Row++, 4}] = {item.ToString()}");
        }
        WriteLog(Receive, $"{new string('-',50)}");

    }

    /// <summary>
    /// 輸出文字檔到被指定的位置並給予指定的檔名。不會重新命名
    /// </summary>
    /// <param name="FullPath">完整的輸出檔案位置及檔名。</param>
    /// <param name="Ecode">以何種編碼寫入文字。</param>
    /// <param name="ExportString">要輸出的文字內容。</param>
    /// <returns>是否有成功輸出檔案。</returns>
    public bool ExportData(string FullPath, Encoding Ecode, string ExportString)
    {
        Boolean ExportStatus = false;
        try
        {
            string[] pathSplit = FullPath.Split('\\');
            string FolderPath = "";
            for (int i = 0; i < pathSplit.Count() - 1; i++) FolderPath += pathSplit[i] + @"\";
            if (!Directory.Exists(FolderPath)) Directory.CreateDirectory(FolderPath);
            if (!File.Exists(FullPath)) File.WriteAllText(FullPath, ExportString + Environment.NewLine, Ecode);
            else File.AppendAllText(FullPath, ExportString + Environment.NewLine, Ecode);
            ExportStatus = true;
        }
        catch
        {
            ExportStatus = false;
        }
        return ExportStatus;
    }

    /// <summary>
    /// 會重新命名
    /// </summary>
    /// <param name="Receive"></param>
    /// <param name="ec"></param>
    /// <param name="Data"></param>
    /// <param name="Ecode"></param>
    public void ExportData(string Receive, ExportConfig ec, int[,] Data, Encoding Ecode)
    {
        string[] TextWay = new string[] { "csv", "txt", "config" };
        string[] ExcelWay = new string[] { "xls", "xlsx" };

        if (ec.LastName != "csv" && Data.GetLength(0) > 100000)
        {
            WriteLog(Receive, "非文字檔輸出，且資料比數超過十萬筆，將副檔名鎖定為csv檔");
            ec.LastName = "csv";
        }

        string FullFileName = ec.ToString();

        string FileDate = Convert.ToString(ec.DTime.Year - 1911) + Convert.ToString(ec.DTime.ToString("MMdd"));
        string FileName = $"{FileDate}_{ec.FileName}_{ec.OpLine}";
        //string ExportFullFileName = CheckAndReName(filePath, FileName, ec.LastName);
        string ExportFullFileName = CheckAndReName(ec.Path1, FullFileName, ec.LastName);
        string FullPath = ec.Path1 + ExportFullFileName;

        if (TextWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"開始文字檔輸出");

            for (int i = 0; i < Data.GetLength(0); i++)
            {
                string sss = "";
                for (int j = 0; j < Data.GetLength(1); j++)
                {
                    sss += $"{Data[i, j]},";
                }
                sss = sss.Substring(0, sss.Length - 1);
                if (!File.Exists(FullPath)) File.WriteAllText(FullPath, sss + Environment.NewLine, Ecode);
                else File.AppendAllText(FullPath, sss + Environment.NewLine, Ecode);
            }
        }
        else if (ExcelWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"開始試算表輸出");
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Data");

            for (int m_Row = 0; m_Row < Data.GetLength(0); m_Row++)
            {
                string sss = "";
                for (int m_Col = 0; m_Col < Data.GetLength(1); m_Col++)
                {
                    ws.Cell(m_Row + 1, m_Col + 1).Style.NumberFormat.Format = "@";
                    ws.Cell(m_Row + 1, m_Col + 1).Value = $"{Data[m_Row, m_Col]}";
                }
            }

            wb.SaveAs(FullPath);
        }
        else
        {
            WriteLog(Receive, $"未發現正確附檔名 -> {ec.LastName}");
        }
    }

    /// <summary>
    /// 會重新命名
    /// </summary>
    /// <param name="Receive"></param>
    /// <param name="ec"></param>
    /// <param name="Data"></param>
    /// <param name="EndCounts"></param>
    /// <param name="Ecode"></param>
    public void ExportData_BigOD(string Receive, ExportConfig ec, string[,] Data, string EndCounts, Encoding Ecode)
    {
        string[] TextWay = new string[] { "csv", "txt", "config" };
        string[] ExcelWay = new string[] { "xls", "xlsx" };

        if (ec.LastName != "csv" && Data.GetLength(0) > 100000)
        {
            WriteLog(Receive, "非文字檔輸出，且資料比數超過十萬筆，將副檔名鎖定為csv檔");
            ec.LastName = "csv";
        }

        string FullFileName = ec.ToString();

        string FileDate = Convert.ToString(ec.DTime.Year - 1911) + Convert.ToString(ec.DTime.ToString("MMdd"));
        string FileName = $"{FileDate}_{ec.FileName}_{ec.OpLine}";
        //string ExportFullFileName = CheckAndReName(filePath, FileName, ec.LastName);
        string ExportFullFileName = CheckAndReName(ec.Path1, FullFileName, ec.LastName);
        string FullPath = ec.Path1 + ExportFullFileName;

        if (TextWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"開始文字檔輸出");

            for (int i = 0; i < Data.GetLength(0); i++)
            {
                string sss = "";
                for (int j = 0; j < Data.GetLength(1); j++)
                {
                    sss += $"{Data[i, j]},";
                }
                sss = sss.Substring(0, sss.Length - 1);
                if (!File.Exists(FullPath)) File.WriteAllText(FullPath, sss + Environment.NewLine, Ecode);
                else File.AppendAllText(FullPath, sss + Environment.NewLine, Ecode);
            }
        }
        else if (ExcelWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"開始試算表輸出");
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Data");

            for (int i = 1; i < Data.GetLength(0); i++)
            {
                ws.Cell(i + 1, Data.GetLength(0) + 1).FormulaR1C1 = $"=SUM(R{i + 1}C2:R{i + 1}C{Data.GetLength(0)})";
                ws.Cell(Data.GetLength(0) + 1, i + 1).FormulaR1C1 = $"=SUM(R2C{i + 1}:R{Data.GetLength(0)}C{i + 1})";
            }
            ws.Cell(1, Data.GetLength(0) + 1).Value = $"出站總量";
            ws.Cell(Data.GetLength(0) + 1, 1).Value = $"進站總量";
            ws.Cell(Data.GetLength(0) + 1, Data.GetLength(0) + 1).FormulaR1C1 = $"" +
                $"=(SUM(AL30:AO30)+SUM(CT30:CW30)+(SUM(DM30:EB30)-DO30))                             " +
                $"+(SUM(Z40:AF40)+SUM(CT40:CX40)+(SUM(DM40:EB40)-DS40))                             " +
                $"+(SUM(Z100:AE100)+SUM(AL100:AO100)+SUM(CT119:CW119)+(SUM(DO100:EB100)-DZ100))" +
                $"+(SUM(AL119:AO119)+(SUM(DM119:EB119)-DO119))                                   " +
                $"+(SUM(B120:EB122)-DP120-DQ121-DR122)                                                          " +
                $"+(SUM(Z123:AF123)+SUM(CT123:CX123)+SUM(DM123:EB123)-DS123-DZ123)                             " +
                $"+(SUM(B124:EB129)-DT124-DU125-DV126-DW127-DX128-DY129)                             " +
                $"+(SUM(Z130:AE130)+SUM(AL130:AO130)+SUM(DO130:EB130)-DZ130)                             " +
                $"+(SUM(B131:EB132)-EA131-EB132)                                                          ";
            //ws.Cell(Data.GetLength(0) + 1, Data.GetLength(0) + 1).Value = $"{EndCounts}";

            for (int m_Row = 0; m_Row < Data.GetLength(0); m_Row++)
            {
                string sss = "";
                for (int m_Col = 0; m_Col < Data.GetLength(1); m_Col++)
                {
                    //ws.Cell(m_Row + 1, m_Col + 1).Style.NumberFormat.Format = "@";
                    ws.Cell(m_Row + 1, m_Col + 1).Value = $"{Data[m_Row, m_Col]}";
                }
            }

            wb.SaveAs(FullPath);
        }
        else
        {
            WriteLog(Receive, $"未發現正確附檔名 -> {ec.LastName}");
        }
    }

    /// <summary>
    /// 會重新命名
    /// </summary>
    /// <param name="Receive"></param>
    /// <param name="ec"></param>
    /// <param name="Data"></param>
    /// <param name="EndCounts"></param>
    /// <param name="Ecode"></param>
    public void ExportData_SmallOD(string Receive, ExportConfig ec, string[,] Data, string EndCounts, Encoding Ecode)
    {
        string[] TextWay = new string[] { "csv", "txt", "config" };
        string[] ExcelWay = new string[] { "xls", "xlsx" };

        if (ec.LastName != "csv" && Data.GetLength(0) > 100000)
        {
            WriteLog(Receive, "非文字檔輸出，且資料比數超過十萬筆，將副檔名鎖定為csv檔");
            ec.LastName = "csv";
        }

        string FullFileName = ec.ToString();

        string FileDate = Convert.ToString(ec.DTime.Year - 1911) + Convert.ToString(ec.DTime.ToString("MMdd"));
        string FileName = $"{FileDate}_{ec.FileName}_{ec.OpLine}";
        //string ExportFullFileName = CheckAndReName(filePath, FileName, ec.LastName);
        string ExportFullFileName = CheckAndReName(ec.Path1, FullFileName, ec.LastName);
        string FullPath = ec.Path1 + ExportFullFileName;

        if (TextWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"開始文字檔輸出");

            for (int i = 0; i < Data.GetLength(0); i++)
            {
                string sss = "";
                for (int j = 0; j < Data.GetLength(1); j++)
                {
                    sss += $"{Data[i, j]},";
                }
                sss = sss.Substring(0, sss.Length - 1);
                if (!File.Exists(FullPath)) File.WriteAllText(FullPath, sss + Environment.NewLine, Ecode);
                else File.AppendAllText(FullPath, sss + Environment.NewLine, Ecode);
            }
        }
        else if (ExcelWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"開始試算表輸出");
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Data");

            for (int i = 1; i < Data.GetLength(0); i++)
            {
                ws.Cell(i + 1, Data.GetLength(0) + 1).FormulaR1C1 = $"=SUM(R{i + 1}C2:R{i + 1}C{Data.GetLength(0)})";
                ws.Cell(Data.GetLength(0) + 1, i + 1).FormulaR1C1 = $"=SUM(R2C{i + 1}:R{Data.GetLength(0)}C{i + 1})";
            }
            ws.Cell(Data.GetLength(0) + 1, Data.GetLength(0) + 1).FormulaR1C1 = $"=SUM(R2C{Data.GetLength(0) + 1}:R{Data.GetLength(0)}C{Data.GetLength(0) + 1})";
            ws.Cell(1, Data.GetLength(0) + 1).Value = $"出站總量";
            ws.Cell(Data.GetLength(0) + 1, 1).Value = $"進站總量";
            ///ws.Cell(Data.GetLength(0) + 1, Data.GetLength(0) + 1).FormulaR1C1 = $"" +
            ///    $"=(SUM(AL30:AO30)+SUM(CT30:CW30)+(SUM(DM30:EC30)-DO30))" +
            ///    $"+(SUM(Z40:AF40)+SUM(CT40:CX40)+(SUM(DM40:EC40)-DS40))" +
            ///    $"+(SUM(Z100:AE100)+SUM(AL100:AO100)+(SUM(DO100:EC100)-DZ100))" +
            ///    $"+(SUM(AL119:AO119)+(SUM(DM119:DP119)-DO119)+EB119+EC119)" +
            ///    $"+(SUM(B120:EC122)-DP120-DQ121-DR122)" +
            ///    $"+(SUM(Z123:AF123)+SUM(CT123:CX123)+SUM(DM123:EC123)-DS123-DZ123)" +
            ///    $"+(SUM(B124:EC129)-DT124-DU125-DV126-DW127-DX128-DY129)" +
            ///    $"+(SUM(Z130:AE130)+SUM(AL130:AO130)+SUM(DO130:EC130)-DZ130)" +
            ///    $"+(SUM(B131:EC132)-EA131-EB132)";
            
            for (int m_Row = 0; m_Row < Data.GetLength(0); m_Row++)
            {
                string sss = "";
                for (int m_Col = 0; m_Col < Data.GetLength(1); m_Col++)
                {
                    //ws.Cell(m_Row + 1, m_Col + 1).Style.NumberFormat.Format = "@";
                    ws.Cell(m_Row + 1, m_Col + 1).Value = $"{Data[m_Row, m_Col]}";
                }
            }

            wb.SaveAs(FullPath);
        }
        else
        {
            WriteLog(Receive, $"未發現正確附檔名 -> {ec.LastName}");
        }
    }

    /// <summary>
    /// 會重新命名
    /// </summary>
    /// <param name="Receive"></param>
    /// <param name="ec"></param>
    /// <param name="DataTitle"></param>
    /// <param name="Data"></param>
    /// <param name="Ecode"></param>
    public void ExportData(string Receive, ExportConfig ec, string DataTitle, int[,] Data, Encoding Ecode)
    {
        string[] TextWay = new string[] { "csv", "txt", "config" };
        string[] ExcelWay = new string[] { "xls", "xlsx" };

        if (ec.LastName != "csv" && Data.GetLength(0) > 100000)
        {
            WriteLog(Receive, "非文字檔輸出，且資料比數超過十萬筆，將副檔名鎖定為csv檔");
            ec.LastName = "csv";
        }

        string FullFileName = ec.ToString();

        string FileDate = Convert.ToString(ec.DTime.Year - 1911) + Convert.ToString(ec.DTime.ToString("MMdd"));
        string FileName = $"{FileDate}_{ec.FileName}_{ec.OpLine}";
        //string ExportFullFileName = CheckAndReName(filePath, FileName, ec.LastName);
        string ExportFullFileName = CheckAndReName(ec.Path1, FullFileName, ec.LastName);
        string FullPath = ec.Path1 + ExportFullFileName;

        if (TextWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"開始文字檔輸出");

            if (DataTitle != "")
            {
                if (!File.Exists(FullPath)) File.WriteAllText(FullPath, DataTitle + Environment.NewLine, Ecode);
                else File.AppendAllText(FullPath, DataTitle + Environment.NewLine, Ecode);
            }

            for (int i = 0; i < Data.GetLength(0); i++)
            {
                string sss = "";
                for (int j = 0; j < Data.GetLength(1); j++)
                {
                    sss += $"{Data[i, j]},";
                }
                sss = sss.Substring(0, sss.Length - 1);
                if (!File.Exists(FullPath)) File.WriteAllText(FullPath, sss + Environment.NewLine, Ecode);
                else File.AppendAllText(FullPath, sss + Environment.NewLine, Ecode);
            }
        }
        else if (ExcelWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"開始試算表輸出");
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Data");

            var dt = DataTitle.Split(',');

            for (int m_Row = 0; m_Row < Data.GetLength(0); m_Row++)
            {
                string sss = "";
                for (int m_Col = 0; m_Col < Data.GetLength(1); m_Col++)
                {
                    if (m_Row == 1)
                    {
                        ws.Cell(m_Row + 1, m_Col + 1).Style.NumberFormat.Format = "@";
                        ws.Cell(m_Row + 1, m_Col + 1).Value = dt[m_Col];
                    }
                    ws.Cell(m_Row + 2, m_Col + 1).Style.NumberFormat.Format = "@";
                    ws.Cell(m_Row + 2, m_Col + 1).Value = $"{Data[m_Row, m_Col]}";
                }
            }

            wb.SaveAs(FullPath);
        }
        else
        {
            WriteLog(Receive, $"未發現正確附檔名 -> {ec.LastName}");
        }
    }

    /// <summary>
    /// 會重新命名
    /// </summary>
    /// <param name="Receive"></param>
    /// <param name="ec"></param>
    /// <param name="DataTitle"></param>
    /// <param name="Data"></param>
    /// <param name="Ecode"></param>
    public void ExportData(string Receive, ExportConfig ec, string DataTitle, string[,] Data, Encoding Ecode)
    {
        string[] TextWay = new string[] { "csv", "txt", "config" };
        string[] ExcelWay = new string[] { "xls", "xlsx" };

        if (ec.LastName != "csv" && Data.GetLength(0) > 100000)
        {
            WriteLog(Receive, "非文字檔輸出，且資料比數超過十萬筆，將副檔名鎖定為csv檔");
            ec.LastName = "csv";
        }

        string FullFileName = ec.ToString();

        ///string FullFileName = ((fileName1.Length > 0) ? $"{fileName1}_" : "") + 
        ///    ((fileName2.Length > 0) ? $"{fileName2}_" : "") +
        ///    ((fileName3.Length > 0) ? $"{fileName3}_" : "") +
        ///    ((fileName4.Length > 0) ? $"{fileName4}_" : "") +
        ///    ec.OpLine;

        string FileDate = Convert.ToString(ec.DTime.Year - 1911) + Convert.ToString(ec.DTime.ToString("MMdd"));
        string FileName = $"{FileDate}_{ec.FileName}_{ec.OpLine}";
        //string ExportFullFileName = CheckAndReName(filePath, FileName, ec.LastName);
        string ExportFullFileName = CheckAndReName(ec.Path1, FullFileName, ec.LastName);
        string FullPath = ec.Path1 + ExportFullFileName;

        if (TextWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"開始文字檔輸出");

            if (DataTitle != "")
            {
                if (!File.Exists(FullPath)) File.WriteAllText(FullPath, DataTitle + Environment.NewLine, Ecode);
                else File.AppendAllText(FullPath, DataTitle + Environment.NewLine, Ecode);
            }

            for (int i = 0; i < Data.GetLength(0); i++)
            {
                string sss = "";
                for (int j = 0; j < Data.GetLength(1); j++)
                {
                    sss += $"{Data[i, j]},";
                }
                sss = sss.Substring(0, sss.Length - 1);
                if (!File.Exists(FullPath)) File.WriteAllText(FullPath, sss + Environment.NewLine, Ecode);
                else File.AppendAllText(FullPath, sss + Environment.NewLine, Ecode);
            }
        }
        else if (ExcelWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"開始試算表輸出");
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Data");

            var dt = DataTitle.Split(',');

            for (int m_Row = 0; m_Row < Data.GetLength(0); m_Row++)
            {
                string sss = "";
                for (int m_Col = 0; m_Col < Data.GetLength(1); m_Col++)
                {
                    if (m_Row + 1 == 1)
                    {
                        ws.Cell(m_Row + 1, m_Col + 1).Style.NumberFormat.Format = "@";
                        ws.Cell(m_Row + 1, m_Col + 1).Value = dt[m_Col];
                    }
                    ws.Cell(m_Row + 2, m_Col + 1).Style.NumberFormat.Format = "@";
                    ws.Cell(m_Row + 2, m_Col + 1).Value = $"{Data[m_Row, m_Col]}";
                }
            }

            wb.SaveAs(FullPath);
        }
        else
        {
            WriteLog(Receive, $"未發現正確附檔名 -> {ec.LastName}");
        }
    }

    /// <summary>
    /// 會重新命名
    /// </summary>
    /// <typeparam name="TResult"></typeparam>
    /// <param name="Receive"></param>
    /// <param name="ec"></param>
    /// <param name="DataTitle"></param>
    /// <param name="Data"></param>
    /// <param name="Ecode"></param>
    public void ExportData<TResult>(string Receive, ExportConfig ec, string DataTitle, List<TResult> Data, Encoding Ecode/*, bool ExportLog = false*/) where TResult : class, new()
    {
        string[] TextWay = new string[] { "csv", "txt", "config" };
        string[] ExcelWay = new string[] { "xls", "xlsx" };

        WriteLog(Receive, $"準備輸出資料筆數: {Data.Count}");
        if (ec.LastName != "csv" && Data.Count > 110000)
        {
            WriteLog(Receive, "非文字檔輸出，且資料比數超過十萬筆，將副檔名鎖定為csv檔");
            ec.LastName = "csv";
        }

        string FullFileName = ec.ToString();

        string FileDate = Convert.ToString(ec.DTime.Year - 1911) + Convert.ToString(ec.DTime.ToString("MMdd"));
        string FileName = $"{FileDate}_{ec.FileName}_{ec.OpLine}";
        //string ExportFullFileName = CheckAndReName(filePath, FileName, ec.LastName);
        string ExportFullFileName = CheckAndReName(ec.Path1, FullFileName, ec.LastName);
        string FullPath = ec.Path1 + ExportFullFileName;

        if (TextWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"開始文字檔輸出");

            if (DataTitle != "")
            {
                if (!File.Exists(FullPath)) File.WriteAllText(FullPath, DataTitle + Environment.NewLine, Ecode);
                else File.AppendAllText(FullPath, DataTitle + Environment.NewLine, Ecode);
            }

            foreach (var item in Data)
            {
                //if (ExportLog) { WriteLog(Receive, $"item: {item.ToString()}"); }
                if (!File.Exists(FullPath)) File.WriteAllText(FullPath, item.ToString() + Environment.NewLine, Ecode);
                else File.AppendAllText(FullPath, item.ToString() + Environment.NewLine, Ecode);
            }
        }
        else if (ExcelWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"開始試算表輸出");
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Data");

            var dt = DataTitle.Split(',');
            int m_Row = 1;

            foreach (var item in Data)
            {
                //if (ExportLog) { WriteLog(Receive, $"item: {item.ToString()}"); }
                var id = item.ToString().Split(',');
                int m_Col = 1;

                foreach (var jtem in id)
                {
                    if (m_Row == 1)
                    {
                        ws.Cell(m_Row, m_Col).Style.NumberFormat.Format = "@";
                        ws.Cell(m_Row, m_Col).Value = dt[m_Col - 1]; 
                    }
                    ws.Cell(m_Row + 1, m_Col).Style.NumberFormat.Format = "@";
                    ws.Cell(m_Row + 1, m_Col).Value = $"{jtem}";
                    m_Col++;
                }
                m_Row++;
            }


            wb.SaveAs(FullPath);
        }
        else
        {
            WriteLog(Receive, $"未發現正確附檔名 -> {ec.LastName}");
        }
    }

    /// <summary>
    /// 會重新命名
    /// </summary>
    /// <typeparam name="TResult"></typeparam>
    /// <param name="Receive"></param>
    /// <param name="ec"></param>
    /// <param name="DataTitle"></param>
    /// <param name="Data"></param>
    /// <param name="Ecode"></param>
    public void ExportData2<TResult>(string Receive, ExportConfig ec, string DataTitle, List<TResult> Data, Encoding Ecode) where TResult : class, new()
    {
        string[] TextWay = new string[] { "csv", "txt", "config" };
        string[] ExcelWay = new string[] { "xls", "xlsx" };

        if (ec.LastName != "csv" && Data.Count > 100000)
        {
            WriteLog(Receive, "非文字檔輸出，且資料比數超過十萬筆，將副檔名鎖定為csv檔");
            ec.LastName = "csv";
        }

        string FullFileName = ec.ToString();

        string FileDate = Convert.ToString(ec.DTime.Year - 1911) + Convert.ToString(ec.DTime.ToString("MMdd"));
        string FileName = $"{FileDate}_{ec.FileName}_{ec.OpLine}";
        //string ExportFullFileName = CheckAndReName(filePath, FileName, ec.LastName);
        string ExportFullFileName = CheckAndReName(ec.Path1, FullFileName, ec.LastName);
        string FullPath = ec.Path1 + ExportFullFileName;

        if (TextWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"開始文字檔輸出");

            if (DataTitle != "")
            {
                if (!File.Exists(FullPath)) File.WriteAllText(FullPath, DataTitle + Environment.NewLine, Ecode);
                else File.AppendAllText(FullPath, DataTitle + Environment.NewLine, Ecode);
            }

            foreach (var item in Data)
            {
                if (!File.Exists(FullPath)) File.WriteAllText(FullPath, item.ToString() + Environment.NewLine, Ecode);
                else File.AppendAllText(FullPath, item.ToString() + Environment.NewLine, Ecode);
            }
        }
        else if (ExcelWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"開始試算表輸出");
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Data");

            var dt = DataTitle.Split(',');
            int m_Row = 1;

            Type type = typeof(TResult);

            foreach (var item in Data)
            {
                for (int m_Col = 1; m_Col <= dt.Length; m_Col++)
                {
                    if (m_Row == 1)
                    {
                        ws.Cell(m_Row, m_Col).Style.NumberFormat.Format = "@";
                        ws.Cell(m_Row, m_Col).Value = dt[m_Col - 1];
                    }
                    ws.Cell(m_Row + 1, m_Col).Style.NumberFormat.Format = "@";
                    ws.Cell(m_Row + 1, m_Col).Value = type.GetProperties()[m_Col - 1].GetValue(item, null);
                }
                m_Row++;
            }


            wb.SaveAs(FullPath);
        }
        else
        {
            WriteLog(Receive, $"未發現正確附檔名 -> {ec.LastName}");
        }
    }

    /// <summary>
    /// 會重新命名
    /// </summary>
    /// <typeparam name="TResult"></typeparam>
    /// <param name="Receive"></param>
    /// <param name="ec"></param>
    /// <param name="DataTitle"></param>
    /// <param name="Data"></param>
    /// <param name="Ecode"></param>
    public void ExportData<TResult>(string Receive, ExportConfig ec, string DataTitle, TResult Data, Encoding Ecode) where TResult : class, new()
    {
        string[] TextWay = new string[] { "csv", "txt", "config" };
        string[] ExcelWay = new string[] { "xls", "xlsx" };

        string FullFileName = ec.ToString();

        //string FileDate = Convert.ToString(ec.DTime.Year - 1911) + Convert.ToString(ec.DTime.ToString("MMdd"));
        //string FileName = $"{FileDate}_{ec.FileName}_{ec.OpLine}";
        //string ExportFullFileName = CheckAndReName(filePath, FileName, ec.LastName);
        string ExportFullFileName = CheckAndReName(ec.Path1, FullFileName, ec.LastName);
        string FullPath = ec.Path1 + ExportFullFileName;

        if (TextWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"開始輸出文字檔");

            if (DataTitle != "")
            {
                if (!File.Exists(FullPath)) File.WriteAllText(FullPath, DataTitle + Environment.NewLine, Ecode);
                else File.AppendAllText(FullPath, DataTitle + Environment.NewLine, Ecode);
            }

            if (!File.Exists(FullPath)) File.WriteAllText(FullPath, Data.ToString() + Environment.NewLine, Ecode);
            else File.AppendAllText(FullPath, Data.ToString() + Environment.NewLine, Ecode);
        }
        else if (ExcelWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"尚未提供試算表輸出");
            //WriteLog(ThisReceive, $"開始輸出試算表");


        }
        else
        {
            WriteLog(Receive, $"未發現正確附檔名 -> {ec.LastName}");
        }
    }

    /// <summary>
    /// 會重新命名
    /// </summary>
    /// <typeparam name="TResult"></typeparam>
    /// <param name="Receive"></param>
    /// <param name="ec"></param>
    /// <param name="DataTitle"></param>
    /// <param name="Data"></param>
    /// <param name="Ecode"></param>
    public void ExportStuctData<TResult>(string Receive, ExportConfig ec, string DataTitle, List<TResult> Data, Encoding Ecode) where TResult : class, new()
    {

        string[] TextWay = new string[] { "csv", "txt", "config" };
        string[] ExcelWay = new string[] { "xls", "xlsx" };

        if (ec.LastName != "csv" && Data.Count > 100000)
        {
            WriteLog(Receive, "非文字檔輸出，且資料比數超過十萬筆，將副檔名鎖定為csv檔");
            ec.LastName = "csv";
        }

        string FullFileName = ec.ToString();

        string FileDate = Convert.ToString(ec.DTime.Year - 1911) + Convert.ToString(ec.DTime.ToString("MMdd"));
        string FileName = $"{FileDate}_{ec.FileName}_{ec.OpLine}";
        //string ExportFullFileName = CheckAndReName(filePath, FileName, ec.LastName);
        string ExportFullFileName = CheckAndReName(ec.Path1, FullFileName, ec.LastName);
        ExportFullFileName = $"{FullFileName}.{ec.LastName}";
        string FullPath = ec.Path1 + ExportFullFileName;

        var sss = ToViewModel<TResult>(Receive, FullPath, DataTitle, 1);
        bool ExportB = true;

        foreach (var item in sss)
        {
            foreach (var jtem in Data)
            {
                if (item.ToString() == jtem.ToString())
                {
                    ExportB = false;
                    break;
                }
            }
            if (!ExportB) break;
        }

        if (!ExportB) return;

        if (TextWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"開始文字檔輸出");

            if (DataTitle != "")
            {
                if (!File.Exists(FullPath)) File.WriteAllText(FullPath, DataTitle + Environment.NewLine, Ecode);
                //else File.AppendAllText(FullPath, DataTitle + Environment.NewLine, Ecode);
            }

            foreach (var item in Data)
            {
                if (!File.Exists(FullPath)) File.WriteAllText(FullPath, item.ToString() + Environment.NewLine, Ecode);
                else File.AppendAllText(FullPath, item.ToString() + Environment.NewLine, Ecode);
            }
        }
        else if (ExcelWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"開始試算表輸出");
            XLWorkbook wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Data");

            var dt = DataTitle.Split(',');
            int m_Row = 1;

            foreach (var item in Data)
            {
                var id = item.ToString().Split(',');
                int m_Col = 1;

                foreach (var jtem in id)
                {
                    if (m_Row == 1)
                    {
                        ws.Cell(m_Row, m_Col).Style.NumberFormat.Format = "@";
                        ws.Cell(m_Row, m_Col).Value = dt[m_Col - 1];
                    }
                    ws.Cell(m_Row + 1, m_Col).Style.NumberFormat.Format = "@";
                    ws.Cell(m_Row + 1, m_Col).Value = $"{jtem}";
                    m_Col++;
                }
                m_Row++;
            }


            wb.SaveAs(FullPath);
        }
        else
        {
            WriteLog(Receive, $"未發現正確附檔名 -> {ec.LastName}");
        }
    }

    /// <summary>
    /// 會重新命名
    /// </summary>
    /// <typeparam name="TResult"></typeparam>
    /// <param name="Receive"></param>
    /// <param name="ec"></param>
    /// <param name="DataTitle"></param>
    /// <param name="Data"></param>
    /// <param name="Ecode"></param>
    /// <param name="NoChecekData"></param>
    public void ExportStuctData<TResult>(string Receive, ExportConfig ec, string DataTitle, TResult Data, Encoding Ecode, ref bool NoChecekData) where TResult : class, new()
    {
        string[] TextWay = new string[] { "csv", "txt", "config" };
        string[] ExcelWay = new string[] { "xls", "xlsx" };

        string filePath = ec.Path1;
        ///string fileName1 = ec.DTime.ToString("yyyyMMdd");
        ///string fileName2 = ec.FirstName;
        ///string fileName3 = ec.FileName;
        ///string fileName4 = ec.SubName;
        ///
        ///string FullFileName = ((fileName1.Length > 0) ? $"{fileName1}_" : "") +
        ///    ((fileName2.Length > 0) ? $"{fileName2}_" : "") +
        ///    ((fileName3.Length > 0) ? $"{fileName3}_" : "") +
        ///    ((fileName4.Length > 0) ? $"{fileName4}_" : "") +
        ///    ec.OpLine;

        string FullFileName = ec.ToString();

        //string FileDate = Convert.ToString(ec.DTime.Year - 1911) + Convert.ToString(ec.DTime.ToString("MMdd"));
        //string FileName = $"{FileDate}_{ec.FileName}_{ec.OpLine}";
        //string ExportFullFileName = CheckAndReName(filePath, FileName, ec.LastName);
        string ExportFullFileName = CheckAndReName(filePath, FullFileName, ec.LastName);
        ExportFullFileName = $"{FullFileName}.{ec.LastName}";
        string FullPath = filePath + ExportFullFileName;

        var sss = ToViewModel<TResult>(Receive, FullPath, DataTitle, 1);
        
        if (!NoChecekData)
        {
            foreach (var item in sss)
            {
                if (item.ToString() == Data.ToString())
                {
                    NoChecekData = false;
                    return;
                }
                else
                {
                    NoChecekData = true;
                }
            }
        }

        if (TextWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"開始輸出文字檔");

            if (DataTitle != "")
            {
                if (!File.Exists(FullPath)) File.WriteAllText(FullPath, DataTitle + Environment.NewLine, Ecode);
                //else File.AppendAllText(FullPath, DataTitle + Environment.NewLine, Ecode);
            }

            if (!File.Exists(FullPath)) File.WriteAllText(FullPath, Data.ToString() + Environment.NewLine, Ecode);
            else File.AppendAllText(FullPath, Data.ToString() + Environment.NewLine, Ecode);
        }
        else if (ExcelWay.Contains(ec.LastName))
        {
            WriteLog(Receive, $"尚未提供試算表輸出");
            //WriteLog(ThisReceive, $"開始輸出試算表");


        }
        else
        {
            WriteLog(Receive, $"未發現正確附檔名 -> {ec.LastName}");
        }
    }

    /// <summary>
    /// 會重新命名
    /// </summary>
    /// <typeparam name="TResult"></typeparam>
    /// <param name="GetThis"></param>
    /// <param name="Ec"></param>
    /// <param name="DataTitle"></param>
    /// <param name="Data"></param>
    /// <param name="GetExport"></param>
    public void ExportVerify<TResult>(ExportConfig Ec, string DataTitle, List<TResult> Data, bool GetExport = false) where TResult : class, new()
    {
        if (GetExport)
        {
            ExecutionMode mMode = ExecutionMode.Debug;
            //GetThis.TestCommonSub("ff");
            string ExportPath = Path_Verify + string.Format(Path_Date, Ec.DTime.ToString("yyyy.MM")) + $@"{Ec.FileName}\";
            string FileDate = Convert.ToString(Ec.DTime.Year - 1911) + Convert.ToString(Ec.DTime.ToString("MMdd"));
            string fileName = string.Format("{0}_{1}({2})_{3}", FileDate, Ec.FileName, "-", Ec.OpLine);
            string ExportCardNum = CheckAndReName(ExportPath, fileName, Ec.LastName);
            string FullPath = ExportPath + ExportCardNum;
            //GetThis.WriteLog(ExecutionMode.Simple, $"輸出檔案名稱:{ExportCardNum}");

            //GetThis.WriteLog(mMode, DataTitle);
            ExportData(FullPath, Encoding.UTF8, DataTitle);
            foreach (var item in Data)
            {
                if (!true)
                {
                    //await Task.Run(() =>
                    //{
                    //    //GetThis.WriteLog(mMode, item.ToString());
                    //    ExportData(FullPath, Encoding.UTF8, item.ToString());
                    //});
                }
                else
                {
                    //GetThis.WriteLog(mMode, item.ToString());
                    ExportData(FullPath, Encoding.UTF8, item.ToString());
                }
            }
        }
    }
    
    /**
    /// <summary>
    /// 將接收到的資料依指定的起始位置依序寫入寫入。
    /// </summary>
    /// <param name="GetThis"></param>
    /// <param name="UseWho">誰需要使用(ex.UseWho = 票卡進出站運量日報表)。</param>
    /// <param name="template">跟UseWho對應給予樣版名稱(ex.template = TicketDayTrafficVolume.xlsx)。</param>
    /// <param name="Data">要寫入Excel的資料。</param>
    /// <param name="TestName">1110203_票卡進出站運量日報表.xlsx。</param>
    /// <returns>是否成功。</returns>
    public Boolean ExportExcel(this Control GetThis, DateTime ClearDate, string UseWho, string template, int[,] Data, string TestName = "")
    {
        Boolean ExportCheck = false;
        XLWorkbook wb = new XLWorkbook();
        //GetThis.WriteLog(ExecutionMode.Simple, "準備輸出檔案");
        int ReportDateStartRow = Convert.ToInt16(Stock_ReportFile.Find(x => x.ReportNameCh == UseWho).ReportDateRow);
        int ReportDateStartColumn = str2num(Stock_ReportFile.Find(x => x.ReportNameCh == UseWho).ReportDateColumn);
        int StartRow = Convert.ToInt16(Stock_ReportFile.Find(x => x.ReportNameCh == UseWho).StartRow);
        int EndRow = Convert.ToInt16(Stock_ReportFile.Find(x => x.ReportNameCh == UseWho).EndRow);
        int StartColumn = str2num(Stock_ReportFile.Find(x => x.ReportNameCh == UseWho).StartColumn);
        int EndColumn = str2num(Stock_ReportFile.Find(x => x.ReportNameCh == UseWho).EndColumn);
        int thisBreakOffRow = 0, thisBreakOffRowNum = 0;
        int thisBreakOffColumn = 0, thisBreakOffColumnNum = 0;
        if (Stock_ReportFile.Find(x => x.ReportNameCh == UseWho).BreakOffRow != "-")
        {
            thisBreakOffRow = Convert.ToInt16(Stock_ReportFile.Find(x => x.ReportNameCh == UseWho).BreakOffRow);
            thisBreakOffRowNum = Convert.ToInt16(Stock_ReportFile.Find(x => x.ReportNameCh == UseWho).BreakOffRowNum);
            //GetThis.WriteLog(ExecutionMode.Simple, "設置行中斷點:" + thisBreakOffRow + " 中斷 " + thisBreakOffRowNum + " 行");
        }
        if (Stock_ReportFile.Find(x => x.ReportNameCh == UseWho).BreakOffColumn != "-")
        {
            thisBreakOffColumn = str2num(Stock_ReportFile.Find(x => x.ReportNameCh == UseWho).BreakOffColumn);
            thisBreakOffColumnNum = Convert.ToInt16(Stock_ReportFile.Find(x => x.ReportNameCh == UseWho).BreakOffColumnNum);
            //GetThis.WriteLog(ExecutionMode.Simple, "設置欄中斷點:" + num2str(thisBreakOffColumn) + " 中斷 " + thisBreakOffRowNum + " 欄");
        }

        //GetThis.WriteLog(ExecutionMode.Simple, "產製日期起始行：" + ReportDateStartRow);
        //GetThis.WriteLog(ExecutionMode.Simple, "產製日期起始欄：" + ReportDateStartColumn);
        //GetThis.WriteLog(ExecutionMode.Simple, "輸出資料起始行：" + StartRow);
        //GetThis.WriteLog(ExecutionMode.Simple, "輸出資料起始欄：" + StartColumn);
        //GetThis.WriteLog(ExecutionMode.Simple, "輸出資料結束行：" + EndRow);
        //GetThis.WriteLog(ExecutionMode.Simple, "輸出資料結束欄：" + EndColumn);
        //GetThis.WriteLog(ExecutionMode.Simple, "資料中斷行    ：" + thisBreakOffRow);
        //GetThis.WriteLog(ExecutionMode.Simple, "資料中斷行數  ：" + thisBreakOffRowNum);
        //GetThis.WriteLog(ExecutionMode.Simple, "資料中斷欄    ：" + thisBreakOffColumn);
        //GetThis.WriteLog(ExecutionMode.Simple, "資料中斷欄數  ：" + thisBreakOffColumnNum);

        //string RePortPath = Path_Program + Path_Config;
        //string TemplatePath = Path_Program + Path_Temp;

        string mYear = Convert.ToString(Convert.ToInt32(ClearDate.ToString("yyyy")) - 1911);
        string mMonth = Convert.ToString(ClearDate.ToString("MM"));
        //string mDay = Convert.ToString(Operation_Start_Date.ToString("dd"));
        string mDay = "";

        //string ReportFile = Stock_ReportFile.Find(x => x.ReportNameCh == UseWho).ReportNameEng;
        string FullName = mYear + mMonth + mDay + "{0}_{1}.xlsx";
        //string FullName = mYear + mMonth + "{0}.xlsx";
        string ReportFileName = string.Format(FullName, "_" + UseWho, Switch_Operation_LineString + TestName);
        //GetThis.WriteLog(ExecutionMode.Simple, "輸出位置      ：" + Path_Program + Path_Report);
        //GetThis.WriteLog(ExecutionMode.Simple, "樣板位置      ：" + Path_Program + Path_Template);
        //GetThis.WriteLog(ExecutionMode.Simple, "輸出檔名      ：" + ReportFileName);
        //GetThis.WriteLog(ExecutionMode.Simple, "樣板檔名      ：" + template);

        Boolean ckeck = CheckFile(Path_Program + Path_Report + ReportFileName);
        if (ckeck)
        {
            //GetThis.WriteLog(ExecutionMode.Simple, "以既有的檔案複寫");
            wb = new XLWorkbook(Path_Program + Path_Report + ReportFileName);
        }
        else
        {
            //GetThis.WriteLog(ExecutionMode.Simple, "以樣板檔案寫入");
            wb = new XLWorkbook(template);
        }

        try
        {
            var ws = wb.Worksheets.First();
            int RowBreak, ColumnBreak;

            //GetThis.WriteLog(ExecutionMode.Simple, "寫入產製日期");
            ws.Cell(ReportDateStartRow, ReportDateStartColumn).Value = ClearDate.ToString("yyyy/MM/01");
            ws.Cell(ReportDateStartRow + 1, ReportDateStartColumn).Value = DateTime.Now.ToString("yyyy/MM/dd");

            if (thisBreakOffRow != 0) RowBreak = thisBreakOffRow - StartRow + 1;
            else RowBreak = Data.GetLength(0);
            if (thisBreakOffColumn != 0) ColumnBreak = thisBreakOffColumn - StartColumn + 1;
            else ColumnBreak = Data.GetLength(1);

            //GetThis.WriteLog(ExecutionMode.Simple, "開始資料寫入");
            for (int RowOffset = 0; RowOffset < RowBreak; RowOffset++)
            {
                try
                {
                    for (int ColumnOffset = 0; ColumnOffset < ColumnBreak; ColumnOffset++)
                    {
                        //GetThis.WriteLog(ExecutionMode.Normal, "sheet.Cell(" + (StartRow + RowOffset) + ", " + (StartColumn + ColumnOffset) + ").Value = " + Data[RowOffset, ColumnOffset]);
                        ws.Cell(StartRow + RowOffset, StartColumn + ColumnOffset).Value = Data[RowOffset, ColumnOffset];
                    }
                    if (thisBreakOffColumn != 0)
                    {
                        for (int ColumnOffset = ColumnBreak; ColumnOffset < EndColumn; ColumnOffset++)
                        {
                            //GetThis.WriteLog(ExecutionMode.Normal, "sheet.Cell(" + (StartRow + RowOffset) + ", " + (StartColumn + thisBreakOffColumnNum + ColumnOffset) + ").Value = " + Data[RowOffset, ColumnOffset]);
                            ws.Cell(StartRow + RowOffset, StartColumn + thisBreakOffColumnNum + ColumnOffset).Value = Data[RowOffset, ColumnOffset];
                        }
                    }
                }
                catch
                {

                }
            }
            //GetThis.WriteLog(ExecutionMode.Simple, "行中斷資料寫入");
            if (thisBreakOffRow != 0)
            {
                for (int RowOffset = RowBreak; RowOffset < EndRow; RowOffset++)
                {
                    try
                    {
                        for (int ColumnOffset = 0; ColumnOffset < ColumnBreak; ColumnOffset++)
                        {
                            //GetThis.WriteLog(ExecutionMode.Normal, "sheet.Cell(" + (StartRow + thisBreakOffRowNum + RowOffset) + ", " + (StartColumn + ColumnOffset) + ").Value = " + Data[StartRow + thisBreakOffRowNum + RowOffset, ColumnOffset]);
                            ws.Cell(StartRow + thisBreakOffRowNum + RowOffset, StartColumn + ColumnOffset).Value = Data[RowOffset, ColumnOffset];
                        }
                        if (thisBreakOffColumn != 0)
                        {
                            for (int ColumnOffset = ColumnBreak; ColumnOffset < EndColumn; ColumnOffset++)
                            {
                                //GetThis.WriteLog(ExecutionMode.Normal, "sheet.Cell(" + (StartRow + thisBreakOffRowNum + RowOffset) + ", " + (StartColumn + thisBreakOffColumnNum + ColumnOffset) + ").Value = " + Data[StartRow + thisBreakOffRowNum + RowOffset, ColumnOffset]);
                                ws.Cell(StartRow + thisBreakOffRowNum + RowOffset, StartColumn + thisBreakOffColumnNum + ColumnOffset).Value = Data[RowOffset, ColumnOffset];
                            }
                        }
                    }
                    catch
                    {

                    }
                }
            }

            //GetThis.WriteLog(ExecutionMode.Simple, "寫入結束，將檔案存成" + ReportFileName);
            ExportCheck = true;
        }
        catch (Exception ex)
        {
            //GetThis.WriteLog(ExecutionMode.Simple, "錯誤！");
            //GetThis.WriteLog(ExecutionMode.Simple, ex.Message);
            MessageBox.Show(ex.Message);
            //return false;
        }
        finally
        {
            wb.SaveAs(Path_Program + Path_Report + ReportFileName);
        }

        return ExportCheck;
    }
    */

    #endregion

    #region Log輸出

    private string _LogLock = "";

    /// <summary>
    /// 輸出Log到主目錄下的Log資料夾。
    /// </summary>
    /// <param name="Receive">哪一個專案使用。</param>
    /// <param name="ExportMode">輸出的層級。</param>
    /// <param name="EventString">要輸出的文字內容。</param>
    public void WriteLog(string Receive, string ExportString)
    {
        lock (_LogLock)
        {
            var dt = DateTime.Now;
            var ss = new StackTrace(true);
            var j = Receive.Split('\\');
            var LogPath = $@"{Path_Program}Log\{dt.ToString("yyyy.MM.dd")}\{((j[0] == "client") ? j[0] : "")}";
            var path = $@"{Path_Program}Log\{dt.ToString("yyyy.MM.dd")}\{Receive}_{dt.ToString("yyyy-MM-dd")}.log";
            var LogMsg = $"{dt:yyyy-MM-dd HH:mm:ss.ff}\t[{ss.GetFrame(1).GetMethod().Name,-30}]\t{ExportString}";
            if (!Directory.Exists(LogPath)) Directory.CreateDirectory(LogPath);
            if (!File.Exists(path)) File.WriteAllText(path, $"===== Creat Log File Success！ ====={Environment.NewLine}");
            File.AppendAllText(path, LogMsg + Environment.NewLine, Encoding.UTF8);
            //     if (ExportMode == ExecutionMode.Simple) Console.ForegroundColor = default;
            //else if (ExportMode <= ExecutionMode.Normal) Console.ForegroundColor = ConsoleColor.Blue;
            //else if (ExportMode == ExecutionMode.Debug) Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine(LogMsg);

        }
    }

    #endregion

    #region 記憶體

    #region 格式化容量大小
    /// <summary>
    /// 格式化容量大小
    /// </summary>
    /// <param name="size">容量（B）</param>
    /// <returns>已格式化的容量</returns>
    private static string FormatSize(double size)
    {
        double d = (double)size;
        int i = 0;
        while ((d > 1024) && (i < 5))
        {
            d /= 1024;
            i++;
        }
        string[] unit = { "B", "KB", "MB", "GB", "TB" };
        return (string.Format("{0} {1}", Math.Round(d, 2), unit[i]));
    }
    #endregion

    #region 獲得記憶體信息API
    [DllImport("kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GlobalMemoryStatusEx(ref MEMORY_INFO mi);

    //定義記憶體的信息結構
    [StructLayout(LayoutKind.Sequential)]
    public struct MEMORY_INFO
    {
        public uint dwLength; //當前結構體大小
        public uint dwMemoryLoad; //當前記憶體使用率
        public ulong ullTotalPhys; //總計物理記憶體大小
        public ulong ullAvailPhys; //可用物理記憶體大小
        public ulong ullTotalPageFile; //總計交換文件大小
        public ulong ullAvailPageFile; //總計交換文件大小
        public ulong ullTotalVirtual; //總計虛擬記憶體大小
        public ulong ullAvailVirtual; //可用虛擬記憶體大小
        public ulong ullAvailExtendedVirtual; //保留 這個值始終為0
    }
    #endregion

    #region 獲得當前記憶體使用情況
    /// <summary>
    /// 獲得當前記憶體使用情況
    /// </summary>
    /// <returns></returns>
    public static MEMORY_INFO GetMemoryStatus()
    {
        MEMORY_INFO mi = new MEMORY_INFO();
        mi.dwLength = (uint)System.Runtime.InteropServices.Marshal.SizeOf(mi);
        GlobalMemoryStatusEx(ref mi);
        return mi;
    }
    #endregion

    public string GetMemory()
    {
        MEMORY_INFO mi = GetMemoryStatus();
        ulong mi_T = mi.ullTotalPhys; // 總記憶體
        ulong mi_C = mi.ullAvailPhys; // 可用
        ulong mi_U = (mi_T - mi_C); // 已用

        var aa = (double)mi_U / (double)mi_T;
        var bb = aa * 100;
        var cc = FormatSize(mi_U);
        var dd = FormatSize(mi_T);
        var c1 = dd.Length;

        string ee = $"     {cc}";
        int ff = ee.Length - c1;
        string gg = ee.Substring(ff, c1);

        var zz = $"{gg} / {dd} ({bb.ToString("#0.00"),6}%)";

        return zz;
    }

    #endregion

    #region 其他

    public bool CheckIP(string Receive, string SearchIP)
    {
        WriteLog(Receive, $"check...");
        IPHostEntry ipEntry = Dns.GetHostEntry(Dns.GetHostName());
        IPAddress[] addr = ipEntry.AddressList;


        for (int i = 0; i < addr.Length; i++)
        {
            WriteLog(Receive, $"IP Address {i}: {addr[i].MapToIPv4().ToString()}");

            string[] sss = addr[i].MapToIPv4().ToString().Split('.');
            string[] yyy = SearchIP.Split('.');
            int cc = 0;

            for (int j = 0; j < yyy.Length; j++)
            {
                if (sss[j] == yyy[j]) { cc++; }
            }

            if (cc == 3) { return true; }
        }

        return false;
    }


    public bool Cmd_(string Receive, string Command, bool UnClose = false)
    {
        try
        {
            WriteLog(Receive, Command);
            using (Process process = new Process())
            {
                process.StartInfo.FileName = "cmd.exe";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardInput = true;
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.RedirectStandardError = true;
                process.StartInfo.CreateNoWindow = true;
                process.Start();
                //process.StandardInput.WriteLine("cls");
                process.StandardInput.WriteLine(Command);
                if (!UnClose)
                {
                    process.StandardInput.WriteLine("exit");
                    string strOutput = process.StandardOutput.ReadToEnd();
                    WriteLog(Receive, $"Show Console Message:\n{strOutput}");
                    process.WaitForExit();
                    process.Close();
                    System.Threading.Thread.Sleep(10 * Time_1s);
                }
            }
            return true;
        }
        catch(Exception ex)
        {

        }



        return false;
    }

    /// <summary>
    /// 刪除指定的程序。
    /// </summary>
    /// <param name="AppProcessName">程序名稱。</param>
    /// <param name="ProcessTitleName">程序標題。(特別判斷：<> -> 不等於空白)</param>
    public bool CheckApplicationAndKill(string AppProcessName, string ProcessTitleName = "")
    {
        bool re = false;
        foreach (System.Diagnostics.Process oItem in System.Diagnostics.Process.GetProcesses())
        {
            if (oItem.ProcessName == AppProcessName)
            {
                bool KillProcess = false;

                // 程式在畫面顯示的名稱
                if (ProcessTitleName == "<>") { if (oItem.MainWindowTitle != "") { KillProcess = true; } } // 自己設定 <> 為非空值
                else if (oItem.MainWindowTitle == ProcessTitleName) { KillProcess = true; } // 其他名稱
                //else if (oItem.MainWindowTitle == $"{ExcelFileName_KKDAY} - Excel") { KillProcess = true;}
                //else if (oItem.MainWindowTitle == $"{ExcelFileName_NTPC} - Excel") { KillProcess = true;}
                if (KillProcess) { oItem.Kill(); re = true; }
            }
        }
        return re;
    }

    public bool CheckApplication(string AppProcessName)
    {
        bool re = false;
        foreach (System.Diagnostics.Process oItem in System.Diagnostics.Process.GetProcesses())
        {
            if (re = oItem.ProcessName == AppProcessName)
            {
                break;
            }
        }
        return re;
    }

    #endregion
}

#region 恩恩

#region this togather

public enum CPUMode
{
    x32 = 0,
    x64 = 1,
    Unknown = -1
}

/// IIS版本的列舉
/// </summary>
[Serializable]
public enum WebServerVersion
{
    /// <summary>
    /// 未知版本
    /// </summary>
    Unknown,
    /// <summary>
    /// IIS 4.0
    /// </summary>
    IIS4,
    /// <summary>
    /// IIS 5.0,5.1
    /// </summary>
    IIS5,
    /// <summary>
    /// IIS 6.0
    /// </summary>
    IIS6,
    /// <summary>
    /// IIS 7.0
    /// </summary>
    IIS7
}

class testCatch_1
{

    /// 取得可用記憶體資訊(MB)
    /// </summary>
    /// <returns></returns>
    ///public static ulong getMemory()
    ///{
    ///    Computer myComputer = new Computer();
    ///    //回傳MB
    ///    return myComputer.Info.AvailablePhysicalMemory / 1024 / 1024;
    ///}

    /// 取得作業系統版本
    /// </summary>
    /// <returns></returns>
    ///public static string checkOSVersion()
    ///{
    ///    Computer myComputer = new Computer();
    ///    string OS_VerStr = string.Format("{0} {1}", myComputer.Info.OSFullName, Environment.OSVersion.ServicePack.ToString());
    ///    return OS_VerStr;
    ///}

    ///public static CPUMode PCCPUMode
    ///{
    ///    get
    ///    {
    ///        // 取得這個執行個體的大小，以位元組為單位。這個屬性的值在 32 位元平台上為 4，而在 64 位元平台上為 8。
    ///        if (IntPtr.Size == 8)
    ///        {
    ///            return CPUMode.x64;
    ///        }
    ///        else
    ///        {
    ///            return CPUMode.x32;
    ///        }
    ///    }
    ///}

    /// 取得.net framework版本
    /// </summary>
    /// <returns></returns>
    ///public static string checkNetVersion()
    ///{
    ///    return Environment.Version.ToString();
    ///}

    /// <summary>
    /// 取得伺服器IIS版本
    /// </summary>
    /// <param name="DomainName"></param>
    /// <returns></returns>
    ///public static WebServerTypes getIISVersion()
    ///{
    ///    string tPath = "IIS://LOCALHOST/W3SVC/INFO";
    ///    DirectoryEntry tEntry = null;
    ///
    ///    try
    ///    {
    ///        tEntry = new DirectoryEntry(tPath);
    ///    }
    ///    catch
    ///    {
    ///        return WebServerVersion.Unknown;
    ///    }
    ///
    ///    int tIISVersion = 5;
    ///    try
    ///    {
    ///        tIISVersion = (int)tEntry.Properties["MajorIISVersionNumber"].Value;
    ///    }
    ///    catch
    ///    {
    ///        return WebServerVersion.Unknown;
    ///    }
    ///
    ///    switch (tIISVersion)
    ///    {
    ///        case 6:
    ///            return WebServerVersion.IIS6;
    ///
    ///        case 7:
    ///            return WebServerVersion.IIS7;
    ///    }
    ///
    ///    //預設給IIS6
    ///    return WebServerVersion.IIS6;
    ///}

}

#endregion

public class testCatch_2 // https://www.itread01.com/article/1501728075.html
{
    public void GetCpuMac()
    {
        ManagementClass mc = new ManagementClass("Win32_Processor");
        ManagementObjectCollection moc = mc.GetInstances();
        string strID = null;
        foreach (ManagementObject mo in moc)
        {
            strID = mo.Properties["ProcessorId"].Value.ToString();
            break;
        }
        Console.WriteLine("CPU ID:" + strID);
    }

    public void GetMotherBordMac()
    {
        ManagementClass mc = new ManagementClass("Win32_BaseBoard");
        ManagementObjectCollection moc = mc.GetInstances();
        string strID = null;
        foreach (ManagementObject mo in moc)
        {
            strID = mo.Properties["SerialNumber"].Value.ToString();
            break;
        }
        Console.WriteLine("主機板 ID:" + strID);
    }

    public void GetHardDiskMac()
    {
        ManagementClass mc = new ManagementClass("Win32_PhysicalMedia");
        //網上有提到，用Win32_DiskDrive，但是用Win32_DiskDrive獲得的硬碟資訊中並不包含SerialNumber屬性。
        ManagementObjectCollection moc = mc.GetInstances();
        string strID = null;
        foreach (ManagementObject mo in moc)
        {
            strID = mo.Properties["SerialNumber"].Value.ToString();
            break;
        }
        Console.WriteLine("硬碟 ID:" + strID);
    }

    public void GetBiosMac()
    {
        ManagementClass mc = new ManagementClass("Win32_BIOS");
        ManagementObjectCollection moc = mc.GetInstances();
        string strID = null;
        foreach (ManagementObject mo in moc)
        {
            strID = mo.Properties["SerialNumber"].Value.ToString();
            break;
        }
        Console.WriteLine("BIOS ID:" + strID);
    }

    public void GetWhatMac()
    {
        ManagementClass mc = new ManagementClass("Win32_Processor");
        ManagementObjectCollection moc = mc.GetInstances();
        foreach (ManagementObject mo in moc)
        {
            Console.WriteLine("\r\n============CUP資訊＝＝＝＝＝＝＝＝＝＝＝");
            foreach (PropertyData pd in mo.Properties)
            {
                Console.WriteLine("\r\n" + pd.Name + "\t");
                if (pd.Value != null)
                {
                    Console.WriteLine(pd.Value.ToString());
                }
            }
            Console.WriteLine("\r\n\r\n============＝＝＝＝＝＝＝＝＝＝＝");
        }
    }
}

public class testCatch_32 // https://learn.microsoft.com/en-us/dotnet/api/system.environment?view=net-7.0
{
    MyCommonds myCommond = new MyCommonds();
    string Recieve = "testCatch2";
    public void GetComputerInfo()
    {
        string ExportStr = "";
        string str;
        //string nl = Environment.NewLine;
        string nl = "";
        //
        //Console.WriteLine();
        ExportStr += ($"\n");
        //Console.WriteLine("-- Environment members --");
        ExportStr += ($"\n-- Environment members --");

        //  Invoke this sample with an arbitrary set of command line arguments.
        //Console.WriteLine("CommandLine               : {0}", Environment.CommandLine);
        ExportStr += ($"\nCommandLine               : {Environment.CommandLine}");

        string[] arguments = Environment.GetCommandLineArgs();
        //Console.WriteLine("GetCommandLineArgs        : {0}", String.Join(", ", arguments));
        ExportStr += ($"\nGetCommandLineArgs        : {String.Join(", ", arguments)}");

        //  <-- Keep this information secure! -->
        //Console.WriteLine("CurrentDirectory          : {0}", Environment.CurrentDirectory);
        ExportStr += ($"\nCurrentDirectory          : {Environment.CurrentDirectory}");

        //Console.WriteLine("ExitCode                  : {0}", Environment.ExitCode);
        ExportStr += ($"\nExitCode                  : {Environment.ExitCode}");

        //Console.WriteLine("HasShutdownStarted        : {0}", Environment.HasShutdownStarted);
        ExportStr += ($"\nHasShutdownStarted        : {Environment.HasShutdownStarted}");

        //  <-- Keep this information secure! -->
        //Console.WriteLine("MachineName               : {0}", Environment.MachineName);
        ExportStr += ($"\nMachineName               : {Environment.MachineName}");

        //Console.WriteLine("NewLine: {0}  first line{0}  second line{0}  third line", Environment.NewLine);
        ExportStr += ($"\nNewLine: {Environment.NewLine}  first line{Environment.NewLine}  second line{Environment.NewLine}  third line");

        //Console.WriteLine("OSVersion                 : {0}", Environment.OSVersion.ToString());
        ExportStr += ($"\nOSVersion                 : {Environment.OSVersion.ToString()}");

        //Console.WriteLine("StackTrace                : {0}'{1}'", Environment.NewLine, Environment.StackTrace);
        ExportStr += ($"\nStackTrace                : {Environment.NewLine}'{Environment.StackTrace}'");

        //  <-- Keep this information secure! -->
        //Console.WriteLine("SystemDirectory           : {0}", Environment.SystemDirectory);
        ExportStr += ($"\nSystemDirectory           : {Environment.SystemDirectory}");

        // 系統累計運行時間
        //Console.WriteLine("TickCount                 : {0}", Environment.TickCount);
        ExportStr += ($"\nTickCount                 : {Environment.TickCount}");

        //  <-- Keep this information secure! -->
        //Console.WriteLine("UserDomainName            : {0}", Environment.UserDomainName);
        ExportStr += ($"\nUserDomainName            : {Environment.UserDomainName}");

        //Console.WriteLine("UserInteractive           : {0}", Environment.UserInteractive);
        ExportStr += ($"\nUserInteractive           : {Environment.UserInteractive}");

        //  <-- Keep this information secure! -->
        //Console.WriteLine("UserName                  : {0}", Environment.UserName);
        ExportStr += ($"\nUserName                  : {Environment.UserName}");

        //Console.WriteLine("Version                   : {0}", Environment.Version.ToString());
        ExportStr += ($"\nVersion                   : {Environment.Version.ToString()}");

        //Console.WriteLine("WorkingSet                : {0}", Environment.WorkingSet);
        ExportStr += ($"\nWorkingSet                : {Environment.WorkingSet}");

        //  No example for Exit(exitCode) because doing so would terminate this example.

        //  <-- Keep this information secure! -->
        string query = "My system drive is        %SystemDrive% and my system root is %SystemRoot%";
        str = Environment.ExpandEnvironmentVariables(query);
        //Console.WriteLine("ExpandEnvironmentVariables: {0}  {1}", nl, str);
        ExportStr += ($"\nExpandEnvironmentVariables: {nl}  {str}");

        //Console.WriteLine("GetEnvironmentVariable    : {0}  My temporary directory is {1}.", nl, Environment.GetEnvironmentVariable("TEMP"));
        ExportStr += ($"\nGetEnvironmentVariable    : {nl}  My temporary directory is {Environment.GetEnvironmentVariable("TEMP")}.");

        //Console.WriteLine("GetEnvironmentVariables   : ");
        ExportStr += ($"\nGetEnvironmentVariables   : ");
        IDictionary environmentVariables = Environment.GetEnvironmentVariables();
        foreach (DictionaryEntry de in environmentVariables)
        {
            Console.WriteLine("  {0,50} = {1}", de.Key, de.Value);
            ExportStr += ($"\n {de.Key,50} = {de.Value}");
        }

        //Console.WriteLine("GetFolderPath             : {0}", Environment.GetFolderPath(Environment.SpecialFolder.System));
        ExportStr += ($"\nGetFolderPath             : {Environment.GetFolderPath(Environment.SpecialFolder.System)}");

        string[] drives = Environment.GetLogicalDrives();
        //Console.WriteLine("GetLogicalDrives          : {0}", String.Join(", ", drives));
        ExportStr += ($"\nGetLogicalDrives          : {String.Join(", ", drives)}");

        myCommond.WriteLog(Recieve, ExportStr);
    }
}

public class testCatch_4_1 // https://www.zendei.com/article/83606.html
{
    static void Main(string[] args)
    {
        Console.WriteLine("總記憶體：" + FormatSize(GetTotalPhys()));
        Console.WriteLine("已使用：" + FormatSize(GetUsedPhys()));
        Console.WriteLine("可使用：" + FormatSize(GetAvailPhys()));
        Console.ReadKey();
    }

    #region 獲得記憶體信息API
    [DllImport("kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GlobalMemoryStatusEx(ref MEMORY_INFO mi);

    //定義記憶體的信息結構
    [StructLayout(LayoutKind.Sequential)]
    public struct MEMORY_INFO
    {
        public uint dwLength; //當前結構體大小
        public uint dwMemoryLoad; //當前記憶體使用率
        public ulong ullTotalPhys; //總計物理記憶體大小
        public ulong ullAvailPhys; //可用物理記憶體大小
        public ulong ullTotalPageFile; //總計交換文件大小
        public ulong ullAvailPageFile; //總計交換文件大小
        public ulong ullTotalVirtual; //總計虛擬記憶體大小
        public ulong ullAvailVirtual; //可用虛擬記憶體大小
        public ulong ullAvailExtendedVirtual; //保留 這個值始終為0
    }
    #endregion

    #region 格式化容量大小
    /// <summary>
    /// 格式化容量大小
    /// </summary>
    /// <param name="size">容量（B）</param>
    /// <returns>已格式化的容量</returns>
    private static string FormatSize(double size)
    {
        double d = (double)size;
        int i = 0;
        while ((d > 1024) && (i < 5))
        {
            d /= 1024;
            i++;
        }
        string[] unit = { "B", "KB", "MB", "GB", "TB" };
        return (string.Format("{0} {1}", Math.Round(d, 2), unit[i]));
    }
    #endregion

    #region 獲得當前記憶體使用情況
    /// <summary>
    /// 獲得當前記憶體使用情況
    /// </summary>
    /// <returns></returns>
    public static MEMORY_INFO GetMemoryStatus()
    {
        MEMORY_INFO mi = new MEMORY_INFO();
        mi.dwLength = (uint)System.Runtime.InteropServices.Marshal.SizeOf(mi);
        GlobalMemoryStatusEx(ref mi);
        return mi;
    }
    #endregion

    #region 獲得當前可用物理記憶體大小
    /// <summary>
    /// 獲得當前可用物理記憶體大小
    /// </summary>
    /// <returns>當前可用物理記憶體（B）</returns>
    public static ulong GetAvailPhys()
    {
        MEMORY_INFO mi = GetMemoryStatus();
        return mi.ullAvailPhys;
    }
    #endregion

    #region 獲得當前已使用的記憶體大小
    /// <summary>
    /// 獲得當前已使用的記憶體大小
    /// </summary>
    /// <returns>已使用的記憶體大小（B）</returns>
    public static ulong GetUsedPhys()
    {
        MEMORY_INFO mi = GetMemoryStatus();
        return (mi.ullTotalPhys - mi.ullAvailPhys);
    }
    #endregion

    #region 獲得當前總計物理記憶體大小
    /// <summary>
    /// 獲得當前總計物理記憶體大小
    /// </summary>
    /// <returns&amp;gt;總計物理記憶體大小（B）&amp;lt;/returns&amp;gt;
    public static ulong GetTotalPhys()
    {
        MEMORY_INFO mi = GetMemoryStatus();
        return mi.ullTotalPhys;
    }
    #endregion
}

public class testCatch_4_2
{
    static void Main(string[] args)
    {
        //需要添加 System.Management 的引用


        //獲取總物理記憶體大小
        ManagementClass cimobject1 = new ManagementClass("Win32_PhysicalMemory");
        ManagementObjectCollection moc1 = cimobject1.GetInstances();
        double available = 0, capacity = 0;
        foreach (ManagementObject mo1 in moc1)
        {
            capacity += ((Math.Round(Int64.Parse(mo1.Properties["Capacity"].Value.ToString()) / 1024 / 1024 / 1024.0, 1)));
        }
        moc1.Dispose();
        cimobject1.Dispose();


        //獲取記憶體可用大小
        ManagementClass cimobject2 = new ManagementClass("Win32_PerfFormattedData_PerfOS_Memory");
        ManagementObjectCollection moc2 = cimobject2.GetInstances();
        foreach (ManagementObject mo2 in moc2)
        {
            available += ((Math.Round(Int64.Parse(mo2.Properties["AvailableMBytes"].Value.ToString()) / 1024.0, 1)));

        }
        moc2.Dispose();
        cimobject2.Dispose();

        Console.WriteLine("總記憶體=" + capacity.ToString() + "G");
        Console.WriteLine("可使用=" + available.ToString() + "G");
        Console.WriteLine("已使用=" + ((capacity - available)).ToString() + "G," + (Math.Round((capacity - available) / capacity * 100, 0)).ToString() + "%");
        Console.ReadKey();
    }
}

public class testCatch_4_3
{
    public static void Main(string[] args)
    {
        PerformanceCounter cpuCounter;
        PerformanceCounter ramCounter;

        cpuCounter = new PerformanceCounter();
        cpuCounter.CategoryName = "Processor";
        cpuCounter.CounterName = "% Processor Time";
        cpuCounter.InstanceName = "_Total";
        cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");
        ramCounter = new PerformanceCounter("Memory", "Available MBytes");



        Console.WriteLine("電腦CPU使用率：" + cpuCounter.NextValue() + "%");
        Console.WriteLine("電腦可使用記憶體：" + ramCounter.NextValue() + "MB");
        Console.WriteLine();



        while (true)
        {
            System.Threading.Thread.Sleep(1000);
            Console.WriteLine("電腦CPU使用率：" + cpuCounter.NextValue() + " %");
            Console.WriteLine("電腦可使用記憶體：" + ramCounter.NextValue() + "MB");
            Console.WriteLine();

            if ((int)cpuCounter.NextValue() > 80)
            {
                System.Threading.Thread.Sleep(1000 * 60);
            }
        }
    }
}















#endregion