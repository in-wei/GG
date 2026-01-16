using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//-
using GlobalCommond.ViewModel;
using UsuallyCommond.MyEnum;
using UsuallyCommond;

namespace Report.Mrt.Check
{
    public partial class ReportCheck1 : Form
    {
        private MyCommonds MyCommond = new MyCommonds();
        private string ThisReceive;
        public List<ReportList> tempReportList = new List<ReportList>();
        List<CheckBoxList> checkboxList;
        public bool IsLife = true;
        bool firstCheck = true;

        public ReportCheck1(bool FirstCheck, List<ReportList> aaa = null)
        {
            InitializeComponent();
            ThisReceive = this.Text;
            firstCheck = FirstCheck;
            RegisterEvents();
            CheckboxUiSet();
            if (firstCheck) AutoChecked();
            else { tempReportList = aaa.ToList(); CheckedCheckBox(); }
        }

        private void RegisterEvents()
        {
            Console.WriteLine($"掛聽各物件");
            //Form
            this.Load += Form_Load;
            this.FormClosing += Form_Close;

            //UI
            checkBox1.CheckedChanged += CheckBox1_CheckedChanged;
            checkBox2.CheckedChanged += CheckBox2_CheckedChanged;
            checkBox3.CheckedChanged += CheckBox3_CheckedChanged;
            checkBox4.CheckedChanged += CheckBox4_CheckedChanged;
            checkBox5.CheckedChanged += CheckBox5_CheckedChanged;
            checkBox6.CheckedChanged += CheckBox6_CheckedChanged;
            checkBox7.CheckedChanged += CheckBox7_CheckedChanged;
            checkBox8.CheckedChanged += CheckBox8_CheckedChanged;
            checkBox9.CheckedChanged += CheckBox9_CheckedChanged;
            checkBox10.CheckedChanged += CheckBox10_CheckedChanged;
            checkBox11.CheckedChanged += CheckBox11_CheckedChanged;
            checkBox12.CheckedChanged += CheckBox12_CheckedChanged;
            checkBox13.CheckedChanged += CheckBox13_CheckedChanged;
            checkBox14.CheckedChanged += CheckBox14_CheckedChanged;
            checkBox15.CheckedChanged += CheckBox15_CheckedChanged;
            checkBox16.CheckedChanged += CheckBox16_CheckedChanged;
            checkBox17.CheckedChanged += CheckBox17_CheckedChanged;
            checkBox18.CheckedChanged += CheckBox18_CheckedChanged;
            checkBox19.CheckedChanged += CheckBox19_CheckedChanged;
            checkBox20.CheckedChanged += CheckBox20_CheckedChanged;

            button1.Click += Button1_Click;
            button2.Click += Button2_Click;
            button3.Click += Button3_Click;
            button4.Click += Button4_Click;
            button5.Click += Button5_Click;

            System.Threading.Thread tt = new System.Threading.Thread(() =>
            {
                while (IsLife)
                {
                    System.Threading.Thread.Sleep(100);
                }
                MyCommond.InvokeIfRequired(this, () =>
                {
                    Hide();
                });
            });
            tt.IsBackground = true;
            tt.Start();

            Console.WriteLine($"完成掛聽");
        }

        #region Form Event

        public void Form_Load(object sender, EventArgs e)
        {
            this.Visible = true;
            Console.WriteLine($"開始變更checkbox項目名稱");
            checkBox1.Text /*--*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)1).Methode.ToDescription()}";
            checkBox2.Text /*--*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)2).Methode.ToDescription()}";
            checkBox3.Text /*--*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)3).Methode.ToDescription()}";
            checkBox4.Text /*--*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)4).Methode.ToDescription()}";
            checkBox5.Text /*--*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)5).Methode.ToDescription()}";
            checkBox6.Text /*--*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)6).Methode.ToDescription()}";
            checkBox7.Text /*--*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)7).Methode.ToDescription()}";
            checkBox8.Text /*--*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)8).Methode.ToDescription()}";
            checkBox9.Text /*--*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)9).Methode.ToDescription()}";
            checkBox10.Text /*-*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)10).Methode.ToDescription()}";
            checkBox11.Text /*-*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)11).Methode.ToDescription()}";
            checkBox12.Text /*-*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)12).Methode.ToDescription()}";
            checkBox13.Text /*-*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)13).Methode.ToDescription()}";
            checkBox14.Text /*-*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)14).Methode.ToDescription()}";
            checkBox15.Text /*-*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)15).Methode.ToDescription()}";
            checkBox16.Text /*-*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)16).Methode.ToDescription()}";
            checkBox17.Text /*-*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)17).Methode.ToDescription()}";
            checkBox18.Text /*-*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)18).Methode.ToDescription()}";
            checkBox19.Text /*-*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)19).Methode.ToDescription()}";
            checkBox20.Text /*-*/ = $"{checkboxList.Find(x => x.Methode == (LrtTxnClientSwitch)20).Methode.ToDescription()}";
            Console.WriteLine($"變更結束");

            
        }

        public void Form_Close(object sender, FormClosingEventArgs e)
        {
            IsLife = false;
            this.Hide();
        }

        #endregion

        private void AutoChecked()
        {
            checkBox1.Checked = true;
            checkBox2.Checked = true;
            checkBox3.Checked = true;
            checkBox4.Checked = true;
            checkBox5.Checked = true;
            checkBox6.Checked = true;
            checkBox7.Checked = true;
            checkBox8.Checked = true;

            //button5.PerformClick();
            this.Hide();
            IsLife = false;
        }

        private void CheckedCheckBox()
        {
            if (tempReportList.Count > 0)
            {
                foreach (var item in tempReportList)
                {
                    if (true)
                    {
                        foreach (var item2 in checkboxList)
                        {
                            if (item2.Methode == item.methode)
                            {
                                item2.CB.Checked = true;
                                break;
                            }
                        }
                    }
                    else
                    {
                        switch ((int)item.methode)
                        {
                            case 0: break;
                            case 1: checkBox1.Checked = true; break;
                            case 2: checkBox2.Checked = true; break;
                            case 3: checkBox3.Checked = true; break;
                            case 4: checkBox4.Checked = true; break;
                            case 5: checkBox5.Checked = true; break;
                            case 6: checkBox6.Checked = true; break;
                            case 7: checkBox7.Checked = true; break;
                            case 8: checkBox8.Checked = true; break;
                            case 9: checkBox9.Checked = true; break;
                            case 10: checkBox10.Checked = true; break;
                            case 11: checkBox11.Checked = true; break;
                            case 12: checkBox12.Checked = true; break;
                            case 13: checkBox13.Checked = true; break;
                            case 14: checkBox14.Checked = true; break;
                            case 15: checkBox15.Checked = true; break;
                            case 16: checkBox16.Checked = true; break;
                            case 17: checkBox17.Checked = true; break;
                            case 18: checkBox18.Checked = true; break;
                            case 19: checkBox19.Checked = true; break;
                            case 20: checkBox20.Checked = true; break;
                            default: break;
                        }
                    }
                }
            }
        }

        #region UI Event - Normal

        private void CheckboxUiSet()
        {
            checkboxList = new List<CheckBoxList>();
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)1, CB = checkBox1 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)2, CB = checkBox2 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)3, CB = checkBox3 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)4, CB = checkBox4 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)5, CB = checkBox5 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)6, CB = checkBox6 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)7, CB = checkBox7 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)8, CB = checkBox8 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)9, CB = checkBox9 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)10, CB = checkBox10 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)11, CB = checkBox11 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)12, CB = checkBox12 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)13, CB = checkBox13 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)14, CB = checkBox14 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)15, CB = checkBox15 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)16, CB = checkBox16 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)17, CB = checkBox17 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)18, CB = checkBox18 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)19, CB = checkBox19 });
            checkboxList.Add(new CheckBoxList() { Methode = (LrtTxnClientSwitch)20, CB = checkBox20 });
        }

        private void A_D_tool(CheckBox cb, LrtTxnClientSwitch methode)
        {
            try
            {
                Console.WriteLine($"{cb.Text} - {((cb.Checked) ? "勾選" : "取消")}");
                if (cb.Checked) 
                {
                    int ss = tempReportList.FindIndex(x => x.methode == methode);
                    if (ss == -1) { tempReportList.Add(new ReportList() { methode = methode }); }
                }
                else 
                {
                    tempReportList.Remove(tempReportList.Find(x => x.methode == methode));
                }
                checkboxList.Find(x => x.Methode == methode).CB = cb;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error!\n{ex.Message}");
            }
            finally
            {

            }
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox1, (LrtTxnClientSwitch)1);
        private void CheckBox2_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox2, (LrtTxnClientSwitch)2);
        private void CheckBox3_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox3, (LrtTxnClientSwitch)3);
        private void CheckBox4_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox4, (LrtTxnClientSwitch)4);
        private void CheckBox5_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox5, (LrtTxnClientSwitch)5);
        private void CheckBox6_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox6, (LrtTxnClientSwitch)6);
        private void CheckBox7_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox7, (LrtTxnClientSwitch)7);
        private void CheckBox8_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox8, (LrtTxnClientSwitch)8);
        private void CheckBox9_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox9, (LrtTxnClientSwitch)9);
        private void CheckBox10_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox10, (LrtTxnClientSwitch)10);
        private void CheckBox11_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox11, (LrtTxnClientSwitch)11);
        private void CheckBox12_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox12, (LrtTxnClientSwitch)12);
        private void CheckBox13_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox13, (LrtTxnClientSwitch)13);
        private void CheckBox14_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox14, (LrtTxnClientSwitch)14);
        private void CheckBox15_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox15, (LrtTxnClientSwitch)15);
        private void CheckBox16_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox16, (LrtTxnClientSwitch)16);
        private void CheckBox17_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox17, (LrtTxnClientSwitch)17);
        private void CheckBox18_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox18, (LrtTxnClientSwitch)18);
        private void CheckBox19_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox19, (LrtTxnClientSwitch)19);
        private void CheckBox20_CheckedChanged(object sender, EventArgs e) => A_D_tool(checkBox20, (LrtTxnClientSwitch)20);

        #endregion

        #region UI Event - Button

        private void Button1_Click(object sender, EventArgs e)
        {
            Console.WriteLine($"日報表全選");
            tempReportList = new List<ReportList>();
            checkBox1.Checked = true;
            checkBox2.Checked = true;
            checkBox3.Checked = true;
            checkBox4.Checked = true;
            checkBox5.Checked = true;
            checkBox6.Checked = true;
            checkBox7.Checked = true;
            checkBox8.Checked = true;
            checkBox9.Checked = false;
            checkBox10.Checked = false;
            checkBox11.Checked = false;
            checkBox12.Checked = false;
            checkBox13.Checked = false;
            checkBox14.Checked = false;
            checkBox15.Checked = false;
            checkBox16.Checked = false;
            checkBox17.Checked = false;
            checkBox18.Checked = false;
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            Console.WriteLine($"重置");
            tempReportList = new List<ReportList>();
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            checkBox7.Checked = false;
            checkBox8.Checked = false;
            checkBox9.Checked = false;
            checkBox10.Checked = false;
            checkBox11.Checked = false;
            checkBox12.Checked = false;
            checkBox13.Checked = false;
            checkBox14.Checked = false;
            checkBox15.Checked = false;
            checkBox16.Checked = false;
            checkBox17.Checked = false;
            checkBox18.Checked = false;
        }

        private void Button3_Click(object sender, EventArgs e)
        {
            Console.WriteLine($"月報表全選");
            tempReportList = new List<ReportList>();
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            checkBox7.Checked = false;
            checkBox8.Checked = false;
            checkBox9.Checked = true;
            checkBox10.Checked = true;
            checkBox11.Checked = true;
            checkBox12.Checked = true;
            checkBox13.Checked = true;
            checkBox14.Checked = true;
            checkBox15.Checked = true;
            checkBox16.Checked = true;
            checkBox17.Checked = true;
            checkBox18.Checked = true;
        }

        private void Button4_Click(object sender, EventArgs e)
        {
            Console.WriteLine($"重置");
            tempReportList = new List<ReportList>();
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox6.Checked = false;
            checkBox7.Checked = false;
            checkBox8.Checked = false;
            checkBox9.Checked = false;
            checkBox10.Checked = false;
            checkBox11.Checked = false;
            checkBox12.Checked = false;
            checkBox13.Checked = false;
            checkBox14.Checked = false;
            checkBox15.Checked = false;
            checkBox16.Checked = false;
            checkBox17.Checked = false;
            checkBox18.Checked = false;
        }

        private void Button5_Click(object sender, EventArgs e)
        {
            Console.WriteLine($"送出");
            IsLife = false;
            this.Hide();
        }

        #endregion

    }
}
