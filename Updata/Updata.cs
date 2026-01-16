using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//--
using System.IO;
using System.Diagnostics;

namespace Updata
{
    public partial class Updata : Form
    {
        MyCommonds MyCommond;
        string ThisReceive = "MainForm";

        private int Time_1s = 1000;

        public Updata()
        {
            InitializeComponent();

            MyCommond = new MyCommonds();
            MyCommond.WriteLog(ThisReceive, "==程式啟動==");
            string ProgramFullPath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName.ToString();
            string ProgramLastWriteTime = File.GetLastWriteTime(ProgramFullPath).ToString();
            MyCommond.WriteLog(ThisReceive, $"This Program Last Write Time： {ProgramLastWriteTime}");

            RegisterEvents();
            Other_Init();

        }

        /// <summary>
        /// 掛聽 From 中的物件。
        /// </summary>
        private void RegisterEvents()
        {
            //Form
            this.Load += Form_Load;

        }

        bool Exe_Auto = false;
        /// <summary>
        /// 其他初始化或是執行程式。
        /// </summary>
        private void Other_Init()
        {
            if (true) MyCommond.CheckApplicationAndKill("EXCEL");

            /// if (true) MyCommond.CheckApplicationAndKill("Main_Console_Normal", "<>");
            /// if (true) MyCommond.CheckApplicationAndKill("Main_Console_Auto", "<>");
            /// if (true) MyCommond.CheckApplicationAndKill("WatchApp", "<>");
            /// 
            /// if (true) MyCommond.CheckApplicationAndKill("Main_Console_Normal", "");
            /// if (true) MyCommond.CheckApplicationAndKill("Main_Console_Auto", "");
            /// if (true) MyCommond.CheckApplicationAndKill("WatchApp", "");

        }

        /// <summary>
        /// 載入主畫面程式。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Form_Load(object sender, EventArgs e)
        {
            /* 開始更新 */ if ( true) { tt2(); }
            /* 開啟程式 */ if (!true) { if (Exe_Auto) MyCommond.Cmd_(ThisReceive, $@"{MyCommond.Path_Program}Main_Console_Auto.exe", true); }
            
            WhileEnd();
        }

        // 改壞掉了
        private void tt()
        {
            string thisPath = MyCommond.Path_Program + @"\";
            string[] search_zip = Directory.GetFiles(thisPath, "Main_Console*.*z*");

            List<string> list_haveFile_exe = new List<string>();
            DirectoryInfo di = new DirectoryInfo($@"{MyCommond.Path_Program}");
            string MyName = this.GetType().Assembly.GetName().Name;

            foreach (var fi in di.GetFiles("*.exe"))
            {
                try
                {
                    if (fi.Name == MyName + ".exe") continue;
                    // List_listviews a = new List_listviews(counts++, fi.Name);
                    list_haveFile_exe.Add(fi.Name);
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, ex.Message);
                }
            }

            MyCommond.CheckFolder(thisPath + @"Other\");

            List<FolderFile2> list_haveFile_exe_2 = new List<FolderFile2>();
            int i = 0;

            foreach (var file in list_haveFile_exe /* haveFile_exe */)
            {
                try
                {
                    // 檔案最後修改時間
                    string ProgramLastWriteTime = File.GetLastWriteTime(file).ToString();
                    string[] s1 = ProgramLastWriteTime.Split(' ');
                    string[] s2 = s1[0].Split('/');
                    DateTime dt = new DateTime(Convert.ToInt16(s2[0]), Convert.ToInt16(s2[1]), Convert.ToInt16(s2[2]));
                    
                    string[] strSplit = file.Split('_');
                    list_haveFile_exe_2.Add(new FolderFile2(++i, file, dt, strSplit[strSplit.Count() - 1]));
                }
                catch (Exception ex)
                {

                }

            }
            if (list_haveFile_exe_2.Count <= 0 || list_haveFile_exe_2 == null)
            {
                MessageBox.Show("未發現檔案!", "警告!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                MyCommond.WriteLog(ThisReceive, "未發現檔案。");
                return;
            }

            List<FolderFile2> List_UpdataFile = (from p in list_haveFile_exe_2 orderby p.FileDate select p).ToList();
            string aa = List_UpdataFile[List_UpdataFile.Count - 1].FileName;
            string[] a1 = aa.Split('\\');
            string FilePath = "";
            string filename = a1[a1.Count() - 1];

            for (int j = 0; j < a1.Length - 1; j++)
            {
                if (a1[j] == "") continue;
                // FilePath += a1[j] + ((j > 0) ? @"\" : "");
                FilePath += a1[j] + @"\";
            }

            MyCommond.Cmd_Decompression(ThisReceive, $@"{thisPath}{aa}");
            MyCommond.MoveFile(thisPath, thisPath + @"\Other", filename);

            MessageBox.Show("更新完成","完成",MessageBoxButtons.OK);
        }

        private void tt2()
        {
            string thisPath = MyCommond.Path_Program + @"\";
            string thisExeName = this.GetType().Assembly.GetName().Name; // 本程式名稱

            var List_exeFile = tttt(thisPath, "*.exe", $"{thisExeName}.exe");
            foreach (var item in List_exeFile)
            {
                if (true)
                {
                    if(MyCommond.CheckApplicationAndKill(item.FileName, "<>"))
                    {
                        MyCommond.CheckApplicationAndKill(item.FileName, "");
                    }
                }
            } // 以工作管理員殺掉這目錄底下除了自己以外正在執行的程式

            var List_updataFile = tttt(thisPath, "NTMC_Volume_Report*.*z*");
            string Updata_FileName = List_updataFile[List_updataFile.Count - 1].FileName; // 取得更新用壓縮檔

            MyCommond.Cmd_Decompression(ThisReceive, Updata_FileName, thisPath); // 解壓縮
            MyCommond.MoveFile(thisPath, $@"{thisPath}Other\", Updata_FileName); // 搬移

            MessageBox.Show("更新完成", "完成", MessageBoxButtons.OK);
        }

        private List<FolderFile2> tttt(string SearchPath, string SearchString, string jump = "")
        {
            List<FolderFile2> Respone = new List<FolderFile2>();

            DirectoryInfo di = new DirectoryInfo(SearchPath);
            int i = 0;
            foreach (var fi in di.GetFiles(SearchString))
            {
                try
                {
                    if (fi.Name == jump) continue;
                    
                    Respone.Add(new FolderFile2(++i, fi));
                }
                catch (Exception ex)
                {
                    MyCommond.WriteLog(ThisReceive, ex.Message);
                }
            }

            return Respone;
        }



        private void WhileEnd()
        {
            MyCommond.WriteLog(ThisReceive, "結束啟動程式");
        StartEnd:
            try
            {
                System.Environment.Exit(0);
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, "重試關閉程序!");
                goto StartEnd;
            }
        }

    }

    public class FolderFile2
    {
        public int No { get; set; }
        public string FileName { get; set; }
        public string FullName { get; set; }
        public string Path { get; set; }
        public DateTime FileDate { get; set; }
        public string FileVersion { get; set; }

        public FolderFile2(int _no, FileInfo fi)
        {
            No = _no;
            FullName = fi.FullName;
            FileName = fi.Name;
            Path = fi.DirectoryName + @"\";
            FileDate = fi.LastWriteTime;


            try
            {
                string[] strSplit1 = fi.Name.Split('_');
                FileVersion = strSplit1[strSplit1.GetLength(0) - 1];
            }
            catch (Exception ex)
            {
                FileVersion = "";
            }
            



        }

        public FolderFile2(int _no = 0, string _filename = "", DateTime _date = new DateTime(), string _version = "")
        {
            No = _no;   /* 流水號 */
            FileName = _filename;  /* 完整陸路徑 */
            FileDate = _date; /* 修改時間 */
            FileVersion = _version; /* 程式的版本 */
        }

        public override string ToString()
        {
            return $"{No,8},{FileName}";
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

        private int Time_1s = 1000;

        private string _LogLock = "";

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
            catch (Exception ex)
            {

            }

            return false;
        }

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
                string CommandString = $@"""C:\Program Files\7-Zip\7z.exe"" x ""{ZipFullPath}""{((overwrite) ? " -aoa" : "")} -o""{ExtractPath}""";
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
    }

}
