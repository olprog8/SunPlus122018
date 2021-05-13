using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Text.RegularExpressions;
using System.Globalization;

using System.Net;
using System.Net.Mail;
using DirectoryServ = System.DirectoryServices;

using Excel = Microsoft.Office.Interop.Excel;
using Directory = System.IO.Directory;
using DirectoryInfo = System.IO.DirectoryInfo;
using File = System.IO.File;
using FileInfo = System.IO.FileInfo;

namespace SunPlus
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        public string gv_winuser;
        public string gv_role;
        public string gv_bunit;

        DAL dal = new DAL();

        private void cbbx_busUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            object selectedItem = cbbx_busUnit.SelectedItem;

            this.lbl_period.Text = dal.GetCurrentPeriod(selectedItem.ToString());
            this.txbx_sunUser.Text = dal.GetSunProfile(System.Environment.UserName.ToString().ToLower())[0];
            this.lbl_limit.Text = dal.GetSunProfile(System.Environment.UserName.ToString())[1];
            this.lbl_delval.Text = dal.JrnalPerDayDeleted(cbbx_busUnit.SelectedItem.ToString(), gv_winuser).ToString();

            gv_role = dal.GetSunProfile(System.Environment.UserName.ToString())[2];
            this.lbl_role.Text = gv_role;

            this.lbl_limRazal.Text = dal.GetSunProfile(System.Environment.UserName.ToString())[3];
            this.lbl_razalVal.Text = dal.RefPerDayRazalloc(cbbx_busUnit.SelectedItem.ToString(), gv_winuser).ToString();

            this.lbl_allocFact.Text = dal.TmPerDay(cbbx_busUnit.SelectedItem.ToString(), gv_winuser).ToString();
            this.btn_allocate.Enabled = true;

            if (this.lbl_period.Text != "none" || this.txbx_sunUser.Text != "none")
            {
                pnl_currentSet.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
                
                if (txbx_journalNumber.Text.Length > 3)
                btn_delete.Enabled = true;
            }

        }

        private void DelJournal_Load(object sender, EventArgs e)
        {
            try
            {
                gv_winuser = System.Environment.UserName.ToString().ToLower();

                this.Text = "  SUN'PLUS  пользователь: " + gv_winuser;

                string[][] emails = new string[8][];
                emails[0] = new string[] { "adm", "akutluev", "Albina.Kutlueva@dhl.com" };
                emails[5] = new string[] { "adm", "olesnits", "Oleg.Lesnitsky@dhl.com" };
                emails[2] = new string[] { "usr2", "ilubenko", "Irina.Lubenko@dhl.com" };
                emails[3] = new string[] { "usr1", "elezhnev", "ekaterina.lezhneva@dhl.com" };
                emails[4] = new string[] { "usr1", "ekurgans", "elena.kurganskaya@dhl.com" };
                emails[1] = new string[] { "usr2", "mschauli", "marina.schaulina@dhl.com" };
                emails[6] = new string[] { "usr2", "kteleev", "Kirill.Teleev@dhl.com" };
                emails[7] = new string[] { "usr2", "scelena", "schelchkova.elena@dhl.com" };

                gv_bunit = dal.GetSunProfile(gv_winuser)[4];

                switch (gv_bunit)
                {
                    case "M11": //Accounting Payble
                    {
                        
                        this.cbbx_busUnit.Items.AddRange(new object[] {
                        "M11","L11","E11","EL1"});
                        this.lbl_Caption.Text = "Журналы";
                        tabControl1.TabPages.Remove(tabPage3);
                        tabControl1.TabPages.Remove(tabPage4);
                        tabControl1.TabPages.Remove(tabPage5);
                        tabControl1.TabPages.Remove(tabPage6);
                        tabControl1.TabPages.Remove(tabPage7);
                        tabControl1.TabPages.Remove(tabPage9);
                        tabControl1.TabPages.Remove(tabPage8);
                        tabControl1.TabPages.Remove(tabPage10);
                        this.pnl_accpay1.Visible = true;
                        this.pnl_accpay2.Visible = true;
                        this.pnl_casher2.Visible = false;

                        break;
                    }

                    case "M13": //Accounting Payble
                    {

                        this.cbbx_busUnit.Items.AddRange(new object[] {
                        "M11","L11","E11","EL1"});
                        this.lbl_Caption.Text = "Журналы";
                        tabControl1.TabPages.Remove(tabPage3);
                        tabControl1.TabPages.Remove(tabPage4);
                        tabControl1.TabPages.Remove(tabPage6);
                        tabControl1.TabPages.Remove(tabPage7);
                        tabControl1.TabPages.Remove(tabPage9);
                        tabControl1.TabPages.Remove(tabPage8);
                        tabControl1.TabPages.Remove(tabPage10);
                        this.pnl_accpay1.Visible = true;
                        this.pnl_accpay2.Visible = true;
                        this.pnl_casher2.Visible = false;

                        break;
                    }

                    case "M12": //Accounting Cashiers
                    {

                        this.cbbx_busUnit.Items.AddRange(new object[] {
                        "M11"});
                        this.lbl_Caption.Text = "Журналы";
                        tabControl1.TabPages.Remove(tabPage2);
                        tabControl1.TabPages.Remove(tabPage3);
                        tabControl1.TabPages.Remove(tabPage5);
                        tabControl1.TabPages.Remove(tabPage6);
                        tabControl1.TabPages.Remove(tabPage7);
                        tabControl1.TabPages.Remove(tabPage9);
                        tabControl1.TabPages.Remove(tabPage8);
                        tabControl1.TabPages.Remove(tabPage10);
                        this.pnl_accpay1.Visible = true;
                        this.pnl_accpay2.Visible = false;
                        this.pnl_casher2.Visible = true;

                        break;
                    }

                    case "M14": //Accounting Banking Group
                    {

                        this.cbbx_busUnit.Items.AddRange(new object[] {
                        "M11","L11"});
                        string year, month, ActPeriod;

                        for (int m = 1; m < 12; m++)
                        {
                            DateTime ActDate = DateTime.Today.AddMonths(-m);

                            year = ActDate.Year.ToString();

                            month = ActDate.Month.ToString();
                            if (month.Length == 1)
                                month = "0" + month;

                            ActPeriod = month + year;
                            //month.Length == 1 ? month = "0" + month : month = month;

                        this.cbbx_TradePeriod.Items.AddRange(new object[] {ActPeriod});
                        }

                        this.cbbx_region.Items.AddRange(new object[] {
                        "MOW","CEN","RFE","SIB","SVR","WES"});
                        this.lbl_Caption.Text = "Журналы";
                        tabControl1.TabPages.Remove(tabPage2);
                        tabControl1.TabPages.Remove(tabPage3);
                        tabControl1.TabPages.Remove(tabPage4);
                        tabControl1.TabPages.Remove(tabPage5);
                        tabControl1.TabPages.Remove(tabPage6);
                        tabControl1.TabPages.Remove(tabPage7);
                        tabControl1.TabPages.Remove(tabPage8);

                        this.pnl_accpay1.Visible = true;
                        this.pnl_accpay2.Visible = false;
                        this.pnl_casher2.Visible = true;

                        break;
                    }

                    case "M15": //Accounting Banking Group (Курганская)
                    {

                        this.cbbx_busUnit.Items.AddRange(new object[] {
                        "M11","L11"});
                        string year, month, ActPeriod;

                        for (int m = 1; m < 12; m++)
                        {
                            DateTime ActDate = DateTime.Today.AddMonths(-m);

                            year = ActDate.Year.ToString();

                            month = ActDate.Month.ToString();
                            if (month.Length == 1)
                                month = "0" + month;

                            ActPeriod = month + year;
                            //month.Length == 1 ? month = "0" + month : month = month;

                            this.cbbx_TradePeriod.Items.AddRange(new object[] { ActPeriod });
                        }

                        this.cbbx_region.Items.AddRange(new object[] {
                        "MOW","CEN","RFE","SIB","SVR","WES"});
                        this.lbl_Caption.Text = "Журналы";
                        tabControl1.TabPages.Remove(tabPage2);
                        tabControl1.TabPages.Remove(tabPage3);
                        tabControl1.TabPages.Remove(tabPage4);
                        tabControl1.TabPages.Remove(tabPage5);
                        tabControl1.TabPages.Remove(tabPage7);
                        tabControl1.TabPages.Remove(tabPage8);

                        this.pnl_accpay1.Visible = true;
                        this.pnl_accpay2.Visible = false;
                        this.pnl_casher2.Visible = true;

                        string esender = "";
                        int i;
                        for (i = 0; i < emails.GetLength(0); i++)
                        {
                            int j;
                            for (j = 1; j < emails[i].GetLength(0); j = j + 3)
                            {
                                if (gv_winuser == emails[i][j])
                                {
                                    esender = emails[i][2].ToLower();
                                    i = emails.GetLength(0);
                                    // MessageBox.Show(esender);
                                    break;
                                }

                            }

                        }
                        this.cbbx_Reciver.Text = esender;
                        this.cbbx_Reciver.Items.AddRange(new object[] {
                        esender});
                        break;
                    }

                    case "A11": //Billing Banking group
                     {
                        this.cbbx_busUnit.Items.AddRange(new object[] {
                        "A11"});

//                    tabPage2.Enabled = false;
                        this.lbl_Caption.Text = "Аллокирование";
                        tabControl1.TabPages.Remove(tabPage1);
                        tabControl1.TabPages.Remove(tabPage2);
                        tabControl1.TabPages.Remove(tabPage4);
                        tabControl1.TabPages.Remove(tabPage5);
                        tabControl1.TabPages.Remove(tabPage6);
                        tabControl1.TabPages.Remove(tabPage7);
                        tabControl1.TabPages.Remove(tabPage9);
                        tabControl1.TabPages.Remove(tabPage8);
                        tabControl1.TabPages.Remove(tabPage10);

                        this.pnl_accpay1.Visible = false;
                        this.pnl_accpay2.Visible = false;
                        this.pnl_casher2.Visible = false;

                        this.btn_allocate.Enabled = false;

                        break;
                      }
                    case "U11": //Admin
                     {
                         this.cbbx_busUnit.Items.AddRange(new object[] {
                        "M11","L11","E11","EL1","A11"});

                         this.cbbx_busUnit.Items.AddRange(new object[] {
                        "M11","L11"});
                         string year, month, ActPeriod;

                         for (int m = 1; m < 12; m++)
                         {
                             DateTime ActDate = DateTime.Today.AddMonths(-m);

                             year = ActDate.Year.ToString();

                             month = ActDate.Month.ToString();
                             if (month.Length == 1)
                                 month = "0" + month;

                             ActPeriod = month + year;
                             //month.Length == 1 ? month = "0" + month : month = month;

                             this.cbbx_TradePeriod.Items.AddRange(new object[] { ActPeriod });
                         }

                         this.cbbx_region.Items.AddRange(new object[] {
                        "MOW","CEN","RFE","SIB","SVR","WES"});

                        int i;
                        for (i = 0; i < emails.GetLength(0); i++)
                        {
                                    this.cbbx_Reciver.Items.AddRange(new object[] {emails[i][2].ToLower()});
                                    // MessageBox.Show(esender);
                        }

                         this.lbl_Caption.Text = "Админ";
                         this.btn_allocate.Enabled = false;
                         this.pnl_accpay1.Visible = true;
                         this.pnl_accpay2.Visible = true;
                         this.pnl_casher2.Visible = true;


                         break;
                     }
                    default: //Default
                     {
                         this.cbbx_busUnit.Items.AddRange(new object[] {
                        "NotFound"});
                         this.lbl_Caption.Text = "";
                         tabControl1.TabPages.Remove(tabPage1);
                         tabControl1.TabPages.Remove(tabPage2);
                         tabControl1.TabPages.Remove(tabPage3);
                         tabControl1.TabPages.Remove(tabPage4);
                         tabControl1.TabPages.Remove(tabPage5);
                         tabControl1.TabPages.Remove(tabPage6);
                         tabControl1.TabPages.Remove(tabPage7);
                         tabControl1.TabPages.Remove(tabPage8);
                         tabControl1.TabPages.Remove(tabPage9);
                         this.pnl_accpay1.Visible = false;
                         this.pnl_accpay2.Visible = false;
                         this.pnl_casher2.Visible = false;

                         break;
                     }
                }

            }
            catch
            {
                gv_winuser = "none";
                this.Text = "  SUN'PLUS  внимание: пользователь не определен ";
            }

            this.txbx_sunUser.Text = "не опред.";
            this.lbl_period.Text = "не опред.";
            this.lbl_limit.Text = "";
            this.lbl_delval.Text = "";
            this.lbl_role.Text = "";
            this.lbl_limRazal.Text = "";
            this.lbl_razalVal.Text = "";
            this.rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + "старт";
            this.txbx_sunUser.Enabled = false;

        }

        private void btn_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_delete_Click(object sender, EventArgs e)
        {
            if (txbx_sunUser.Text == "не опред." || lbl_period.Text == "не опред.")
            {
                MessageBox.Show(" Проверьте наличие значения в поле Бизнес Юнит! ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (txbx_journalNumber.Text.ToString() == string.Empty || txbx_journalNumber.Text.ToString() == "")
            {
                MessageBox.Show(" Введите номер журнала! ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            int cntJournalDeleted = dal.JrnalPerDayDeleted(cbbx_busUnit.SelectedItem.ToString(), gv_winuser);

            if (int.Parse(lbl_limit.Text) == cntJournalDeleted && gv_role != "R05")
//            if (int.Parse(lbl_limit.Text) == cntJournalDeleted)
            {
                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Удаление журнала для " + cbbx_busUnit.SelectedItem.ToString() + " невозможно, превышение лимита на удаление за текущий день;\n" + rtxbx_info.Text;
            }
            else
            {
                btn_delete.Enabled = false;
                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + dal.journalDelLayout(cbbx_busUnit.SelectedItem.ToString(), txbx_journalNumber.Text.ToString(), lbl_period.Text, txbx_sunUser.Text, lbl_role.Text, gv_winuser) + "\n" + rtxbx_info.Text;
                btn_delete.Enabled = true;
                this.lbl_delval.Text = dal.JrnalPerDayDeleted(cbbx_busUnit.SelectedItem.ToString(), gv_winuser).ToString();
            }

        }

        private void txbx_journalNumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar < 48 || e.KeyChar > 58 && e.KeyChar !=8)
                e.Handled = true;
        }



        //private void btn_realloc_Click(object sender, EventArgs e)
        //{
        //    string ref1 = "", ref2 = "", ref3 = "", ref4 = "", ref5 = "", refs = "";

        //    if (txbx_journalRealloc.Text != string.Empty && txbx_journalRealloc.Text.ToString().Length != 0)
        //    {
        //        if (txbx_ref1.Text.ToString().Trim().Length > 0)
        //        {
        //            ref1 = txbx_ref1.Text.ToString().Trim();
        //            refs = refs + ref1;

        //            if (txbx_ref2.Text.ToString().Trim().Length > 0)
        //            {
        //                ref1 = txbx_ref2.Text.ToString().Trim();
        //                refs = refs + ", " + ref2;
        //            }
        //            if (txbx_ref3.Text.ToString().Trim().Length > 0)
        //            {
        //                ref1 = txbx_ref3.Text.ToString().Trim();
        //                refs = refs + ", " + ref3;
        //            }
        //            if (txbx_ref4.Text.ToString().Trim().Length > 0)
        //            {
        //                ref1 = txbx_ref4.Text.ToString().Trim();
        //                refs = refs + ", " + ref4;
        //            }
        //            if (txbx_ref5.Text.ToString().Trim().Length > 0)
        //            {
        //                ref1 = txbx_ref5.Text.ToString().Trim();
        //                refs = refs + ", " + ref5;
        //            }
        //        }
        //     }

        //        //DAL.JournalReallocation(txbx_journalRealloc.Text.ToString().Length, );
        //}


        private void txbx_journalNumber_TextChanged(object sender, EventArgs e)
        {
            if (this.txbx_journalNumber.Text.Length > 2)
                btn_delete.Enabled = true;
        }

        /*
        private void btn_TransMatching_Click(object sender, EventArgs e)
        {
            this.rtxbx_info.Text = dal.TransMatching();
         * 
                 *             if (e.KeyChar < 48 || e.KeyChar > 58 && e.KeyChar !=8)
                e.Handled = true;

         * *             if (this.txbx_journalNumber.Text.Length > 2)
                btn_delete.Enabled = true;
         * 
        }
         */

        private void btn_razalloc_Click(object sender, EventArgs e)
        {
            if (txbx_sunUser.Text == "не опред." || lbl_period.Text == "не опред.")
            {
                MessageBox.Show(" Проверьте наличие значения в поле Бизнес Юнит! ", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (txbx_journalNumber2.Text.ToString() == string.Empty || txbx_journalNumber2.Text.ToString() == "")
            {
                MessageBox.Show(" Введите номер журнала! ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (txbx_reference.Text.ToString() == string.Empty || txbx_reference.Text.ToString() == "")
            {
                MessageBox.Show(" Введите Референс! ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            int cntRefRazalloc = dal.RefPerDayRazalloc(cbbx_busUnit.SelectedItem.ToString(), gv_winuser);


            if (int.Parse(lbl_limRazal.Text) == cntRefRazalloc && gv_role != "R05")
            {
                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Разаллокирование референса для " + cbbx_busUnit.SelectedItem.ToString() + " невозможно, превышение лимита разаллокирований за текущий день;\n" + rtxbx_info.Text;
            }
            else
            {
                btn_razalloc.Enabled = false;
                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + dal.RazAllocLayout(cbbx_busUnit.SelectedItem.ToString(), txbx_journalNumber2.Text.ToString(), txbx_reference.Text.ToString(), txbx_sunUser.Text, lbl_role.Text, gv_winuser) + "\n" + rtxbx_info.Text;
//                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + dal.RazAllocLayout(cbbx_busUnit.SelectedItem.ToString(), txbx_journalNumber.Text.ToString(), txbx_reference.Text.ToString()) + "\n" + rtxbx_info.Text;
                btn_razalloc.Enabled = true;
                this.lbl_razalVal.Text = dal.RefPerDayRazalloc(cbbx_busUnit.SelectedItem.ToString(), gv_winuser).ToString();
            }

        }

        private void txbx_journalNumber2_TextChanged(object sender, EventArgs e)
        {
            if (this.txbx_journalNumber2.Text.Length > 3 && this.txbx_reference.Text.Length > 2)
                this.btn_razalloc.Enabled = true;
        }

        private void txbx_reference_TextChanged(object sender, EventArgs e)
        {
            if (this.txbx_journalNumber2.Text.Length > 3 && this.txbx_reference.Text.Length > 2)
                this.btn_razalloc.Enabled = true;
        }

//Аллокирование
        private void btn_allocate_Click(object sender, EventArgs e)
        {
            if (txbx_sunUser.Text == "не опред." || lbl_period.Text == "не опред.")
            {
                MessageBox.Show(" Проверьте наличие значения в поле Бизнес Юнит! ", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            //Определяем сколько раз запускали ТМ
            int cntTm = dal.TmPerDay(cbbx_busUnit.SelectedItem.ToString(), gv_winuser);


            if (cntTm >= int.Parse(this.lbl_limAlloc.Text) && gv_role != "R05")
            {
                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Аллокирование для " + cbbx_busUnit.SelectedItem.ToString() + " невозможно, превышение лимита Аллокирования за текущий день;\n" + rtxbx_info.Text;
            }
            else
            {
                this.btn_allocate.Enabled = false;
                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + dal.TmAllocLayout(cbbx_busUnit.SelectedItem.ToString(), txbx_sunUser.Text, lbl_role.Text, gv_winuser) + "\n" + rtxbx_info.Text;
                //                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + dal.RazAllocLayout(cbbx_busUnit.SelectedItem.ToString(), txbx_journalNumber.Text.ToString(), txbx_reference.Text.ToString()) + "\n" + rtxbx_info.Text;
                this.btn_allocate.Enabled = true;
                this.lbl_allocFact.Text = dal.TmPerDay(cbbx_busUnit.SelectedItem.ToString(), gv_winuser).ToString();
            }

        }

//Добавление курьера
        static int FindText(string sentence, string pattern)
        {

            int result;
            result = 0;

            //pattern = @"^[A-Z]{1,20}\s[A-Z]{1,20}\s[A-Z]{1,20}";
            //pattern = @"[А-Я]{1,4}\s[А-Я]{1,4}";

            Regex newReg = new Regex(pattern);
            MatchCollection matches = newReg.Matches(sentence);

            if (matches.Count == 1)
            {
                result = 1;
            }

            return result;
        }

        private void btn_addCourier_Click(object sender, EventArgs e)
        {
            //SendMail();
            
            if (txbx_sunUser.Text == "не опред." || lbl_period.Text == "не опред.")
            {
                MessageBox.Show(" Проверьте наличие значения в поле Бизнес Юнит! ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (FindText(txbx_fiocourier.Text, @"[А-Я]{1,20}\s[А-Я]{1,20}\s[А-Я]{1,20}") != 1)
            {
                MessageBox.Show(" Курьер не добавлен. Проверьте формат ввода ФИО курьера [Фамилия Имя Отчество - через пробел] ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "]  Курьер <" + this.txbx_fiocourier.Text + "> не добавлен. Проверьте формат ввода ФИО курьера;\n" + rtxbx_info.Text;
                return;
            }


            int cntCourierIns = dal.CourierPerDayIns(cbbx_busUnit.SelectedItem.ToString(), gv_winuser);

            if (int.Parse(lbl_limit.Text) == cntCourierIns && gv_role != "R05")
            //            if (int.Parse(lbl_limit.Text) == cntJournalDeleted)
            {
                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Добавление Курьера для " + cbbx_busUnit.SelectedItem.ToString() + " невозможно, превышен лимит на добавление за текущий день;\n" + rtxbx_info.Text;
            }
            else
            {
                btn_addCourier.Enabled = false;

                string ins_res = dal.courierInsLayout(cbbx_busUnit.SelectedItem.ToString(), txbx_fiocourier.Text.ToString(), lbl_period.Text, txbx_sunUser.Text, lbl_role.Text, gv_winuser);
                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + ins_res + "\n" + rtxbx_info.Text;

                btn_addCourier.Enabled = true;
                this.lbl_adedCourier.Text = dal.CourierPerDayIns(cbbx_busUnit.SelectedItem.ToString(), gv_winuser).ToString();

                if (ins_res.IndexOf("не добавлен") == 0)
                {
                    this.txbx_fiocourier.Text = "";
                }
            }

        }


        private void btn_unblockCur_Click(object sender, EventArgs e)
        {
            if (txbx_sunUser.Text == "не опред." || lbl_period.Text == "не опред.")
            {
                MessageBox.Show(" Проверьте наличие значения в поле Бизнес Юнит! ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (FindText(txbx_fiocourier.Text, @"[А-Я]{1,20}\s[А-Я]{1,20}\s[А-Я]{1,20}") != 1)
            {
                MessageBox.Show(" Курьер не разблокирован. Проверьте формат ввода ФИО курьера [Фамилия Имя Отчество - через пробел] ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "]  Курьер <" + this.txbx_fiocourier.Text + "> не разблокирован. Проверьте формат ввода ФИО курьера;\n" + rtxbx_info.Text;
                return;
            }

            btn_addCourier.Enabled = false;
            rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + dal.courierUnblockLayout(cbbx_busUnit.SelectedItem.ToString(), txbx_fiocourier.Text.ToString(), lbl_period.Text, txbx_sunUser.Text, lbl_role.Text, gv_winuser) + "\n" + rtxbx_info.Text;
            btn_addCourier.Enabled = true;
            btn_unblockCur.Enabled = true;
            this.txbx_fiocourier.Text = "";
        }


//Вывод списка Файлов

        private void btn_upNFilToExcel_Click(object sender, EventArgs e)
        {

            Excel.Application xlApp = new Excel.Application();
            string[] FileNames = new[]{""};

            try
            {

                xlApp.Workbooks.Add(Type.Missing);
                xlApp.Interactive = false;
                xlApp.EnableEvents = false;

                Excel.Worksheet xlSheet;
                xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                xlSheet.Name = "Список Файлов";

                FileNames = GetFiles(this.txbx_pathForList.Text);

                if (FileNames.Length == 0)
                {
                    MessageBox.Show("Файлов не обнаружено");
                    return;
                }

                int colInd = 1;
                int rowInd = 1;

                //C:\OL2\dpo_management.xls
                for (rowInd = 1; rowInd < FileNames.Length; rowInd++)
                {
                    xlApp.Cells[rowInd, colInd] = ProcessFName(FileNames[rowInd - 1].ToString());
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                this.rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Из папки <" + this.txbx_pathForList.Text + "> в Excel выгружено " + (FileNames.Length - 1) + " имен(и) файлов. " + "\n" + rtxbx_info.Text;
                MessageBox.Show("Из папки <" + this.txbx_pathForList.Text + "> в Excel выгружено " + (FileNames.Length-1) + " имен(и) файлов. ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                
                xlApp.Visible = true;

                xlApp.Interactive = true;
                xlApp.ScreenUpdating = true;
                xlApp.UserControl = true;
            }
            


        }


        //Вывод списка Каталогов
        private void btn_upNDirToExcel_Click(object sender, EventArgs e)
        {

            Excel.Application xlApp = new Excel.Application();
            string[] DirNames = new[] { "" };

            try
            {

                xlApp.Workbooks.Add(Type.Missing);
                xlApp.Interactive = false;
                xlApp.EnableEvents = false;

                Excel.Worksheet xlSheet;
                xlSheet = (Excel.Worksheet)xlApp.Sheets[1];
                xlSheet.Name = "Список Каталогов";

                DirNames = GetDirs(this.txbx_pathForList.Text);

                if (DirNames.Length == 0)
                {
                    MessageBox.Show("Каталогов не обнаружено");
                    return;
                }

                int colInd = 1;
                int rowInd = 1;

                //C:\OL2\dpo_management.xls
                for (rowInd = 1; rowInd < DirNames.Length; rowInd++)
                {
                    xlApp.Cells[rowInd, colInd] = ProcessFName(DirNames[rowInd - 1].ToString());
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                this.rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Из папки <" + this.txbx_pathForList.Text + "> в Excel выгружено " + (DirNames.Length-1) + " имен(и) каталогов. " + "\n" + rtxbx_info.Text;
                MessageBox.Show("Из папки <" + this.txbx_pathForList.Text + "> в Excel выгружено " + (DirNames.Length-1) + " имен(и) Каталогов. ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                xlApp.Visible = true;

                xlApp.Interactive = true;
                xlApp.ScreenUpdating = true;
                xlApp.UserControl = true;
            }

        }

        private string[] GetFiles(string path)
        {
            string[] files = Directory.GetFiles(path);
            DirectoryInfo di = new DirectoryInfo(path);
            return files;
        }

        private string[] GetFiles2(string path)
        {
            FileInfo[] fi;
            ArrayList al = new ArrayList();

            DirectoryInfo di = new DirectoryInfo(path);
            fi = di.GetFiles("*.xls");
            foreach (FileInfo f in fi)
            {
                al.Add(f.Name);
            }
            string[] files = (string[])al.ToArray(typeof(string));
            return files;
        }

        private string[] GetDirs(string path)
        {
            string[] Dirs = Directory.GetDirectories(path);
            return Dirs;
        }

        private string ProcessFName(string FullFPath)
        {
            int pos = 0;

            pos = FullFPath.IndexOf('\\');

            while (pos != -1)
            {
                FullFPath = FullFPath.Substring(pos + 1);
                pos = FullFPath.IndexOf('\\');
            }

            return FullFPath;
        }


        void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        static void SendMail(string reciever, string mesBody, string busUnit, bool OpenPer)
        {
            //smtp сервер
            string smtpHost = "smtp.prg-dc.dhl.com";
            //smtp порт
            int smtpPort = 25;
            //логин
            string login = "ruhrpadm";
            //пароль
            string pass = "Rh121212";

            //создаем подключение
            SmtpClient client = new SmtpClient(smtpHost, smtpPort);
            client.Credentials = new NetworkCredential(login, pass);

            //От кого письмо
            string from = "ruhrpadm@dhl.com";
            //Кому письмо
            //string to = "oleg.lesnitsky@dhl.com";
            //Тема письма
            string subject;
            subject = "";
            if (busUnit == "M11")
            subject = "Инфо: ПЕРИОД в 1C Trade ZAO ";
            else if (busUnit == "E11")
            subject = "Инфо: ПЕРИОД в 1C Trade OOO ";

            if (OpenPer)
                subject = subject + " Открыт[+]. ";
            else
                subject = subject + " Закрыт[-]. ";

            //Текст письма
            //string body = "Привет! \n\n\n Это тестовое письмо от C Sharp";

            //Создаем сообщение
            MailMessage mess = new MailMessage(from, reciever, subject, mesBody);
            mess.BodyEncoding = Encoding.UTF8;
            mess.CC.Add("oleg.lesnitsky@dhl.com");
            mess.CC.Add("albina.kutlueva@dhl.com");

//            MailMessage mess2 = new MailMessage(from, "oleg.lesnitsky@dhl.com", subject, mesBody);
//            mess2.BodyEncoding = Encoding.UTF8;
            try
            {
                client.Send(mess);
//                client.Send(mess2);
                //Console.WriteLine("Message send");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.ReadKey();
            }
        }

        static void SendMail2(string reciever, string mesBody)
        {
            //smtp сервер
            string smtpHost = "smtp.prg-dc.dhl.com";
            //smtp порт
            int smtpPort = 25;
            //логин
            string login = "ruhrpadm";
            //пароль
            string pass = "Rh121212";

            //создаем подключение
            SmtpClient client = new SmtpClient(smtpHost, smtpPort);
            client.Credentials = new NetworkCredential(login, pass);

            //От кого письмо
            string from = "ruhrpadm@dhl.com";
            //Кому письмо
            //string to = "oleg.lesnitsky@dhl.com";
            //Тема письма
            string subject = "Смена пароля в SUN Accounting ";

            //Текст письма
            //string body = "Привет! \n\n\n Это тестовое письмо от C Sharp";

            //Создаем сообщение
            MailMessage mess = new MailMessage(from, reciever, subject, mesBody);
            mess.BodyEncoding = Encoding.UTF8;
            mess.CC.Add("oleg.lesnitsky@dhl.com");
            mess.CC.Add("albina.kutlueva@dhl.com");

            //            MailMessage mess2 = new MailMessage(from, "oleg.lesnitsky@dhl.com", subject, mesBody);
            //            mess2.BodyEncoding = Encoding.UTF8;
            try
            {
                //
                client.Send(mess);
                //                client.Send(mess2);
                //                Console.WriteLine("Message send");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.ReadKey();
            }

        }

        private void txbx_fiocourier_TextChanged(object sender, EventArgs e)
        {
            if (this.txbx_fiocourier.Text.Length > 10)
                btn_addCourier.Enabled = true;
                btn_unblockCur.Enabled = true;
        }

        private void btn_openPer_Click(object sender, EventArgs e)
        {
            object busUnit = cbbx_busUnit.SelectedItem;
            if (busUnit == null)
            {
                MessageBox.Show(" Укажите значения в поле БизнесЮнит [M11 - Trade ЗАО]:", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (busUnit.ToString() != "M11" && busUnit.ToString() != "E11")
            {
                MessageBox.Show(" Укажите значения в поле БизнесЮнит [M11 - Trade ЗАО]:", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }


            if (this.cbbx_TradePeriod.SelectedItem.ToString().Length != 6 || this.cbbx_region.SelectedItem.ToString().Length != 3 || this.txbx_ComntPeriod.Text.Length < 5 || this.cbbx_Reciver.Text.ToString().Length <= 16)
            {
                MessageBox.Show(" Проверьте наличие значения в полях Открыть Период:, Регион:, Комментарий:, Электронный адрес :", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            lbl_GetCurPerd_DoubleClick(this.btn_openPer, new EventArgs());

            if (this.lbl_CurTradePerd.Text == this.cbbx_TradePeriod.SelectedItem.ToString())
            {
                        rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + "Выбранный Период открыт." + "\n" + rtxbx_info.Text;
                        return;
            }

            rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + dal.OpenTradePerdLayout(this.cbbx_region.SelectedItem.ToString(), this.cbbx_TradePeriod.Text.ToString(), this.txbx_ComntPeriod.Text.ToString(), "TRZ", "M14", gv_winuser) + "\n" + rtxbx_info.Text;

            lbl_GetCurPerd_DoubleClick(this.btn_openPer, new EventArgs());
            //Подключиться к базе и поменять период по параметрам (период, регион)
            //Записать в лог (кто, во-сколько, период, регион, причина
            //Отправить оповещение заинтересованным пользователям
            SendMail(this.cbbx_Reciver.Text.ToString(), "Коллеги, добрый день!\n\nПериод: " + this.cbbx_TradePeriod.Text.ToString() + " \nдля региона: " + this.cbbx_region.SelectedItem.ToString() + " \nпо причине: <" + this.txbx_ComntPeriod.Text.ToString() + "> \nоткрыт.\nНе забудьте Закрыть период!\n\nЭто служебное сообщение (и адрес эл.почты), просьба на него не отвечать. \nСпасибо. \n\nС Уважением, Олег Лесницкий.\nmailto:oleg.lesnitsky@dhl.com", this.cbbx_busUnit.SelectedItem.ToString(), true);

        }

        private void lbl_GetCurPerd_DoubleClick(object sender, EventArgs e)
        {
            if (this.cbbx_region.SelectedItem == null || this.cbbx_region.SelectedItem.ToString().Length != 3)
            {
                MessageBox.Show(" Проверьте наличие значения в полях Период с:, Регион:", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            this.lbl_CurTradePerd.Text = dal.GetOpenTradePerd(this.cbbx_region.SelectedItem.ToString(), "M14", gv_winuser);

        }

        private void cbbx_region_SelectedIndexChanged(object sender, EventArgs e)
        {
            lbl_GetCurPerd_DoubleClick(this.cbbx_region, new EventArgs());
        }

        private void btn_closePer_Click(object sender, EventArgs e)
        {
            object busUnit = cbbx_busUnit.SelectedItem;
            if (busUnit == null)
            {
                MessageBox.Show(" Укажите значения в поле БизнесЮнит [M11 - Trade ЗАО]:", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (busUnit.ToString() != "M11" && busUnit.ToString() != "E11")
            {
                MessageBox.Show(" Укажите значения в поле БизнесЮнит [M11 - Trade ЗАО]:", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (this.cbbx_region.SelectedItem.ToString().Length != 3)
            {
                MessageBox.Show(" Проверьте наличие значения в поле Регион:", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            lbl_GetCurPerd_DoubleClick(this.cbbx_region, new EventArgs());

            rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + dal.OpenTradePerdLayout(this.cbbx_region.SelectedItem.ToString(), "TRZ", "M14", gv_winuser) + "\n" + rtxbx_info.Text;


            lbl_GetCurPerd_DoubleClick(this.cbbx_region, new EventArgs());

            SendMail(this.cbbx_Reciver.Text.ToString(), "Коллеги, добрый день!\n\nПериод для региона: " + this.cbbx_region.SelectedItem.ToString() + "\nзакрыт.\n\nЭто служебное сообщение (и адрес эл.почты), просьба на него не отвечать.\n\nС Уважением, Олег Лесницкий.\nmailto:oleg.lesnitsky@dhl.com", this.cbbx_busUnit.SelectedItem.ToString(), false);

        }

        private void btn_rename_Click(object sender, EventArgs e)
        {
            string[] FileNames = new[]{""};

            try
            {

                FileNames = GetFiles(this.txbx_pathForList.Text);

                if (FileNames.Length == 0)
                {
                    MessageBox.Show("Файлов не обнаружено");
                    return;
                }

                int fileInd = 1;
                FileInfo fl;

                //C:\OL2\dpo_management.xls
                for (fileInd = 0; fileInd < FileNames.Length; fileInd++)
                {
                    fl = new FileInfo(FileNames[fileInd]);
                    fl.MoveTo(this.txbx_pathForList.Text + "\\" + fl.Name.Substring(0,8));

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            MessageBox.Show("Finish");
        }


        public void ADcon()
        {
            //DirectoryEntry userEntry = new DirectoryEntry("LDAP://prg-dc.dhl.com/CN=olesnits,OU=RU,OU=Users,OU=RUMOW,DC=prg-dc,DC=dhl,DC=Com", "olesnits", "Dhl@702red)7");
            DirectoryServ.DirectoryEntry userEntry = new DirectoryServ.DirectoryEntry("LDAP://rootDSE", "olesnits", "Dhl@702red)7");
            DirectoryServ.PropertyCollection props = userEntry.Properties;
            foreach (string prop in props.PropertyNames)
            {
                DirectoryServ.PropertyValueCollection values = props[prop];
                foreach (string val in values)
                {
                    Console.Write(prop + ": ");
                    Console.WriteLine(val);
                }
            }
            
        }

        private void cbbx_Reciver_SelectedValueChanged(object sender, EventArgs e)
        {
            this.cbbx_Reciver.Text = this.cbbx_Reciver.SelectedItem.ToString();
        }

        private void cbbx_Reciver_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.cbbx_Reciver.Text = this.cbbx_Reciver.SelectedItem.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string curdate;
            DateTime TheDate = DateTime.Today.AddDays(-5);
            curdate = TheDate.Year.ToString() + "-" + TheDate.Month.ToString() + "-" + TheDate.Day.ToString();
            MessageBox.Show(curdate);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string[] FileNames = new[] { "" };
            Excel.Workbook ObjWorkBook = null;


            try
            {

                FileNames = GetFiles(this.txbx_pathForList.Text);

                if (FileNames.Length == 0)
                {
                    MessageBox.Show("Файлов не обнаружено");
                    return;
                }

                int fileInd = 1;

                //C:\OL2\dpo_management.xls
                for (fileInd = 0; fileInd < FileNames.Length; fileInd++)
                {

                    //fl = new FileInfo(FileNames[fileInd]);
                    //fl.MoveTo(this.txbx_pathForList.Text + "\\" + fl.Name.Substring(0, 8));

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                ObjWorkBook.Close(false, null, null);
            }


            MessageBox.Show("Finish");
        }

        private void cbtn_convert_Click(object sender, EventArgs e)
        {
            string[] FileNames = new[] { "" };
            Excel.Workbook ObjWorkBook = null;
            Excel.Worksheet ObjWorkSheet = null;
            Excel.Range ObjWorkRange = null;

            DirectoryInfo df = new DirectoryInfo(this.txbx_pathForListSF.Text+"\\xls");
            if (df.Exists)
            {
             MessageBox.Show("Папка XLS уже существует в папке <" + this.txbx_pathForListSF.Text + ">. Переименуйте её для создания новой.");
             return;
            }
            else
            {
                df.Create();
            }
            try
            {

                FileNames = GetFiles2(this.txbx_pathForListSF.Text);

                if (FileNames.Length == 0)
                {
                    MessageBox.Show("Файлы не обнаружены");
                    return;
                }

                Excel.Application XlsApp = new Excel.Application();

                int fileInd = 1;
                string newFileName = "";

                //C:\OL2\dpo_management.xls
                for (fileInd = 0; fileInd < FileNames.Length; fileInd++)
                {
//                    ObjWorkBook = XlsApp.Workbooks.Open(this.txbx_pathForListSF.Text + "\\" + FileNames[fileInd]);
                    ObjWorkBook = XlsApp.Workbooks.Open("Y:\\OL\\Convert\\Budget1.xlsx");
                    ObjWorkSheet = (Excel.Worksheet)XlsApp.Sheets[1];
                    ObjWorkRange = (Excel.Range)XlsApp.Range["B19:M19"];
                    ObjWorkRange.Merge();

                    newFileName = df.FullName + "\\" + FileNames[fileInd].Substring(0, FileNames[fileInd].Length - 4) + "m.xls";
                    ObjWorkBook.SaveAs(newFileName, Excel.XlFileFormat.xlExcel7, null, null, null, null,
                    Excel.XlSaveAsAccessMode.xlExclusive, null, null, null);
                    ObjWorkBook.Close(false, newFileName, null);

                    //fl = new FileInfo(FileNames[fileInd]);
                    //fl.MoveTo(this.txbx_pathForList.Text + "\\" + fl.Name.Substring(0, 8));
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
                ObjWorkBook.Close(false, null, null);
                MessageBox.Show("XML Файлы обработаны!");
        }

        private void cbtn_sendInfo_Click_1(object sender, EventArgs e)
        {

            string mess2;

            string[] users = new String[] {"irina.donets@dhl.com","Ирина","Id121212",
"tatiyana.kozlovskaya@dhl.com","Татьяна","Tk121212",
"elena.kurganskaya@dhl.com","Елена","Ek121212",
"ekaterina.lezhneva@dhl.com","Екатерина","El121212",
"nataliya.rudakovskaya@dhl.com","Наталья","Nr121212",
"anzhelika.rudenko@dhl.com","Анжелика","Ar121212",
"lyana.salpagarova@dhl.com","Ляна","Ls121212",
"anastasia.nesterova@dhl.com","Анастасия","An121212"
 };

            if (users.Length % 3 > 0)
            {
                MessageBox.Show("Некорректные исходные данные", "Скрипт остановлен!");
                return;
            }

            for(int i=0;i<users.Length-2;i=i+3)
            {
            //MessageBox.Show(users[i].ToString()+"; "+users[i+1].ToString()+"; "+users[i+2].ToString());

            mess2 = "Добрый день, " + users[i + 1].ToString() + "!\n\n";
            mess2 = mess2 + "В связи с правилами CRISP для системы Sun Systems меняется политика паролей.\n";
            mess2 = mess2 + "Ваш изменен пароль (с заглавной буквы): " + users[i + 2].ToString() + ".\n\n";

            mess2 = mess2 + "Данный пароль временный, просьба сменить его функцией смены пароля CPA.\n";
            mess2 = mess2 + "Пароль должен быть не менее 8 символов.\n\nС Уважением, Олег Лесницкий.\nmailto:oleg.lesnitsky@dhl.com";

            SendMail2(users[i].ToString(), mess2);

            }

            MessageBox.Show("Сообщения отправлены!");

//Добрый день, Наташа!

//В связи с правилами CRISP для системы Sun Systems меняется политика паролей.
//Для пользователя NFO изменен пароль (с заглавной буквы): Nf121212.

//Данный пароль временный, просьба сменить его функцией смены пароля CPA.
//Пароль должен быть не менее 8 символов.


        }

        private void btn_loadRateToSun_Click(object sender, EventArgs e)
        {

            object busUnit = cbbx_busUnit.SelectedItem;
            if (busUnit == null || busUnit.ToString() != "M11")
            {
                MessageBox.Show(" Укажите значения в поле БизнесЮнит [M11 - Trade ЗАО]:", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            DateTime rate_datetime;
            rate_datetime = DateTime.Today.AddDays(1);


            if (this.dtpk_dateRate.Text != null)
            rate_datetime = DateTime.Parse(this.dtpk_dateRate.Text.ToString(), CultureInfo.CreateSpecificCulture("ru-RU"));


            if (rate_datetime > DateTime.Today || rate_datetime < DateTime.Today.AddDays(-31))
            {
                MessageBox.Show(" Значение Даты Курсов должно быть \n не позднее текущей даты и не более 31 дня ранее...\n Проверьте значение даты!", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            this.rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + dal.LoadRateToSun(rate_datetime, gv_winuser, this.txbx_sunUser.Text) +"\n" + rtxbx_info.Text;

        }

        private void txbx_journaUselNumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar < 48 || e.KeyChar > 58 && e.KeyChar !=8)
            e.Handled = true;
        }

        private void txbx_journaUselNumber_TextChanged(object sender, EventArgs e)
        {
            if (this.txbx_journaUselNumber.Text.Length > 5)
                this.btn_offUse.Enabled = true;
        }

        private void btn_offUse_Click(object sender, EventArgs e)
        {

            object busUnit = cbbx_busUnit.SelectedItem;
            if (busUnit == null || busUnit.ToString() != "M11")
            {
                MessageBox.Show(" Укажите значения в поле БизнесЮнит [M11 - Trade ЗАО]:", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            btn_offUse.Enabled = false;
            this.rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + dal.JournalInUseOff(cbbx_busUnit.SelectedItem.ToString(), this.txbx_journaUselNumber.Text.ToString(), lbl_period.Text, txbx_sunUser.Text, lbl_role.Text, gv_winuser) + "\n" + rtxbx_info.Text;
            btn_offUse.Enabled = true;
            //this.lbl_delval.Text = dal.JrnalPerDayDeleted(cbbx_busUnit.SelectedItem.ToString(), gv_winuser).ToString();

        }


        /*
        //Смена кодировки
        mess.SubjectEncoding = Encoding.Default;
        mess.BodyEncoding = Encoding.Default;
        mess.Headers["Content-type"] = "text/plain; charset=windows-1251";
        //Вложение для письма
        //Если нужно не одно вложение, для каждого создаем отдельный Attachment
        Attachment attData = new Attachment(@"D:\att.zip");
        //прикрепляем вложение
        mess.Attachments.Add(attData);
         */

        /*
         
                     

            if (int.Parse(lbl_limit.Text) == cntJournalDeleted && gv_role != "R05")
//            if (int.Parse(lbl_limit.Text) == cntJournalDeleted)
            {
                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Удаление журнала для " + cbbx_busUnit.SelectedItem.ToString() + " невозможно, превышение лимита на удаление за текущий день;\n" + rtxbx_info.Text;
            }
            else
            {
                btn_delete.Enabled = false;
                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + dal.journalDelLayout(cbbx_busUnit.SelectedItem.ToString(), txbx_journalNumber.Text.ToString(), lbl_period.Text, txbx_sunUser.Text, lbl_role.Text, gv_winuser) + "\n" + rtxbx_info.Text;
                btn_delete.Enabled = true;
                this.lbl_delval.Text = dal.JrnalPerDayDeleted(cbbx_busUnit.SelectedItem.ToString(), gv_winuser).ToString();
            }

        }

        private void txbx_journalNumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar < 48 || e.KeyChar > 58 && e.KeyChar !=8)
                e.Handled = true;
        }

         
             */
            ////проверка на количество разаллокирований на текущий день по этому пользователю
            ////проверка что данный журнал и референс имеются и относятся к данному пользователю и текущему периоду
            ////проверка что в данном журнале+референсе есть съаллокированные записи
            ////разаллокация (стандартный скрипт)
            ////проверка что в данном журнале+референсе нет съаллокированных записей
            ////логгирование разаллокации
            ////вывод сообщения - журнал разаллокирован
    }
}
