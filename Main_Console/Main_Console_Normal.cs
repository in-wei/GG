using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Main_Console_Normal
{
    public partial class Form1_Main_Normal : Form
    {
        MyCommonds MyCommond;
        string ThisReceive = "MainForm";

        public Form1_Main_Normal()
        {
            MyCommond = new MyCommonds();
            MyCommond.WriteLog(ThisReceive, $"Main Start");
            InitializeComponent();
            //notifyIcon1.Visible = true;
            //notifyIcon1.ShowBalloonTip(1, "0", "0", ToolTipIcon.None);
            this.ShowInTaskbar = false;
            this.WindowState = FormWindowState.Minimized;
            this.Hide();
            try
            {
                if ( true) MyCommond.CheckApplicationAndKill("EXCEL");
                MyCommond.WriteLog(ThisReceive, "輸入將啟動的程式:");
            Start_1:
                string x = Console.ReadLine();

                if (x.Equals("QUIT", StringComparison.CurrentCultureIgnoreCase) || x.Equals("Q", StringComparison.CurrentCultureIgnoreCase))
                { WhileEnd(); }
                else if (x.Equals("LRT", StringComparison.CurrentCultureIgnoreCase))
                { MyCommond.WriteLog(ThisReceive, $"您輸入{x,4}，準備啟動輕軌報表程式。"); try { Report.Lrt.UI.Form_Lrt form_Lrt = new Report.Lrt.UI.Form_Lrt(); form_Lrt.Show(); } catch { } }
                else if (x.Equals("MRT", StringComparison.CurrentCultureIgnoreCase))
                { MyCommond.WriteLog(ThisReceive, $"您輸入{x,4}，準備啟動捷運報表程式。"); try { Report.Mrt.UI.Form_Mrt form_Mrt = new Report.Mrt.UI.Form_Mrt(); form_Mrt.Show(); } catch { } }
                else if (x.Equals("test", StringComparison.CurrentCultureIgnoreCase))
                { MyCommond.WriteLog(ThisReceive, $"您輸入{x,4}，準備啟動捷運資料庫程式。"); try { Test.Form1 form_Mrt = new Test.Form1(); form_Mrt.Show(); } catch { } }
                else if (x.Equals("auto", StringComparison.CurrentCultureIgnoreCase))
                { MyCommond.WriteLog(ThisReceive, $"您輸入{x,4}，準備啟動捷運自動程式。"); try { Main_Console_Auto.Form1_Main_Auto from_Auto = new Main_Console_Auto.Form1_Main_Auto(); from_Auto.Show(); } catch { } }
                else if (x.Equals("QR", StringComparison.CurrentCultureIgnoreCase))
                { MyCommond.WriteLog(ThisReceive, $"您輸入{x,4}，準備啟動測試報表程式。"); try { QR_Code_FindOneDayTicket.UI.Form1_QR form_qr = new QR_Code_FindOneDayTicket.UI.Form1_QR(); form_qr.Show(); } catch { } }
                else
                {
                    MyCommond.WriteLog(ThisReceive, $"您輸入的{x}，未被列入啟動關鍵字");
                    MyCommond.WriteLog(ThisReceive, "請重新輸入:");
                    goto Start_1;
                }
            }
            catch (Exception ex)
            {
                MyCommond.WriteLog(ThisReceive, $"ERROR!\n{ex.Message}");
            }
        }

        private void WhileEnd()
        {
            Console.WriteLine("press any key to exit the process...");
            Console.ReadKey(false);
            //while (Console.ReadKey().Key != ConsoleKey.Enter) { }
            MyCommond.WriteLog(ThisReceive, "結束啟動程式");
            System.Environment.Exit(0);
        }

        private void Form1_Main_Load(object sender, EventArgs e)
        {

        }

        private void Form1_Main_Normal_Load(object sender, EventArgs e)
        {

        }
    }
}
