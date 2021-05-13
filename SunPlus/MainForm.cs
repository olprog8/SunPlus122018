using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;

using System.Threading; // For setting the Localization of the thread to fit
using System.Globalization; // the of the MS Excel localization, because of the MS bug

using System.Net;
using System.Net.Mail;
using DirectoryServ = System.DirectoryServices;

using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
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
        public string t_period;

        DAL dal = new DAL();

        public static string[][] Emails
        {
            get
            {
                //ВАЖНО!!! не пропускать индексы и при изменении количества изменять размерность массива
                string[][] emails = new string[10][];
                emails[0] = new string[] { "adm", "akutluev", "Albina.Kutlueva@dhl.com" };
                emails[4] = new string[] { "adm", "olesnits", "Oleg.Lesnitsky@dhl.com" };
                emails[1] = new string[] { "usr2", "ilubenko", "Irina.Lubenko@dhl.com" };
                emails[2] = new string[] { "usr1", "elezhnev", "ekaterina.lezhneva@dhl.com" };
                emails[3] = new string[] { "usr1", "ekurgans", "elena.kurganskaya@dhl.com" };
                emails[5] = new string[] { "usr2", "kteleev", "Kirill.Teleev@dhl.com" };

                emails[6] = new string[] { "usr2", "emogurov", "Evgeniya.Mogurova@dhl.ru" };
                emails[7] = new string[] { "usr3", "opyatykh", "Olga.Pyatykh@dhl.com" };
                emails[8] = new string[] { "usr3", "nbiserov", "Nataliya.Biserova@dhl.ru" };
                emails[9] = new string[] { "usr3", "ekryzhan", "elena.kirillova@dhl.ru" };

                return emails;
            }
        }

        public static string GetSender(string winuser)
        {
            string sender = "";

            for (int i = 0; i < Emails.GetLength(0); i++)
            {
                int j;
                for (j = 1; j < Emails[i].GetLength(0); j = j + 3)
                {
                    if (winuser == Emails[i][j])
                    {
                        sender = Emails[i][2].ToLower();
                        i = Emails.GetLength(0);
                        // MessageBox.Show(esender);
                        break;
                    }

                }

            }

            return sender;
        }

        private void cbbx_busUnit_SelectedIndexChanged(object sender, EventArgs e)
        {
            object selectedItem = cbbx_busUnit.SelectedItem;
            //
            string curPer = dal.GetCurrentPeriod(selectedItem.ToString().Trim());
            //string curPer = "2017009";
            this.lbl_period.Text = curPer;

            this.txbx_sunUser.Text = dal.GetSunProfile(System.Environment.UserName.ToString().ToLower())[0];
            this.lbl_limit.Text = dal.GetSunProfile(System.Environment.UserName.ToString())[1];
            //this.lbl_delval.Text = dal.JrnalPerDayModif(cbbx_busUnit.SelectedItem.ToString(), gv_winuser,1).ToString();
            this.lbl_delval.Text = "0";
            //this.lbl_perShift.Text = dal.JrnalPerDayModif(cbbx_busUnit.SelectedItem.ToString(), gv_winuser, 2).ToString();
            this.lbl_perShift.Text = "0";

            gv_role = dal.GetSunProfile(System.Environment.UserName.ToString())[2];
            this.lbl_role.Text = gv_role;

            this.lbl_limRazal.Text = dal.GetSunProfile(System.Environment.UserName.ToString())[3];
            //this.lbl_razalVal.Text = dal.RefPerDayRazalloc(cbbx_busUnit.SelectedItem.ToString(), gv_winuser).ToString();
            this.lbl_razalVal.Text = "0";

            //this.lbl_allocFact.Text = dal.TmPerDay(cbbx_busUnit.SelectedItem.ToString(), gv_winuser).ToString();
            this.lbl_allocFact.Text = "0";
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
                this.lbl_ProtocolConds.Text = "За текущий день";

                string[][] emails = MainForm.Emails;

                gv_bunit = dal.GetSunProfile(gv_winuser)[4];
                //                gv_bunit = "M16";

                switch (gv_bunit)
                {
                    case "M11": //Accounting Payble
                        {

                            this.cbbx_busUnit.Items.AddRange(new object[] {
//                        "M15","L15","E15","EL5"});
                        "RU0","RU1","RU4","RU5"});
                            this.lbl_Caption.Text = "Журналы";
                            //                        tabControl1.TabPages.Remove(tabPage1);
                            TabCollection.TabPages.Remove(tabPage3);
                            TabCollection.TabPages.Remove(tabPage4);
                            TabCollection.TabPages.Remove(tabPage5);
                            TabCollection.TabPages.Remove(tabPage6);
                            TabCollection.TabPages.Remove(tabPage7);
                            TabCollection.TabPages.Remove(tabPage9);
                            TabCollection.TabPages.Remove(tabPage8);
                            TabCollection.TabPages.Remove(tabPage10);
                            TabCollection.TabPages.Remove(tabPage11);
                            TabCollection.TabPages.Remove(tabPage12);
                            TabCollection.TabPages.Remove(tabPage13);
                            TabCollection.TabPages.Remove(tabPage14);
                            TabCollection.TabPages.Remove(tabPage15);
                            TabCollection.TabPages.Remove(tabPage17);
                            this.pnl_accpay1.Visible = true;
                            this.pnl_accpay2.Visible = true;
                            this.pnl_casher2.Visible = false;
                            this.pnl_shift4.Visible = false;

                            TabCollection.TabPages.Remove(tabPage1);
                            TabCollection.TabPages.Remove(tabPage2);


                            break;
                        }

                    case "M13": //Taxation
                        {

                            this.cbbx_busUnit.Items.AddRange(new object[] {
                        "RU0","RU1","RU4","RU5"});
                            this.lbl_Caption.Text = "Журналы";
                            TabCollection.TabPages.Remove(tabPage1);
                            TabCollection.TabPages.Remove(tabPage3);
                            TabCollection.TabPages.Remove(tabPage4);
                            TabCollection.TabPages.Remove(tabPage6);
                            TabCollection.TabPages.Remove(tabPage7);
                            TabCollection.TabPages.Remove(tabPage9);
                            TabCollection.TabPages.Remove(tabPage8);
                            TabCollection.TabPages.Remove(tabPage10);
                            TabCollection.TabPages.Remove(tabPage11);
                            TabCollection.TabPages.Remove(tabPage12);
                            TabCollection.TabPages.Remove(tabPage13);
                            TabCollection.TabPages.Remove(tabPage14);
                            TabCollection.TabPages.Remove(tabPage17);
                            this.pnl_accpay1.Visible = true;
                            this.pnl_accpay2.Visible = true;
                            this.pnl_casher2.Visible = false;
                            this.pnl_shift4.Visible = false;

                            TabCollection.TabPages.Remove(tabPage2);

                            break;
                        }

                    case "M12": //Accounting Cashiers
                        {

                            this.cbbx_busUnit.Items.AddRange(new object[] {
                        "RU0"});
                            this.lbl_Caption.Text = "Журналы";
                            TabCollection.TabPages.Remove(tabPage2);
                            TabCollection.TabPages.Remove(tabPage3);
                            TabCollection.TabPages.Remove(tabPage5);
                            TabCollection.TabPages.Remove(tabPage6);
                            TabCollection.TabPages.Remove(tabPage7);
                            TabCollection.TabPages.Remove(tabPage9);
                            TabCollection.TabPages.Remove(tabPage8);
                            TabCollection.TabPages.Remove(tabPage10);
                            TabCollection.TabPages.Remove(tabPage11);
                            TabCollection.TabPages.Remove(tabPage12);
                            TabCollection.TabPages.Remove(tabPage13);
                            TabCollection.TabPages.Remove(tabPage14);
                            TabCollection.TabPages.Remove(tabPage15);
                            TabCollection.TabPages.Remove(tabPage17);
                            this.pnl_accpay1.Visible = true;
                            this.pnl_accpay2.Visible = false;
                            this.pnl_casher2.Visible = true;
                            this.pnl_shift4.Visible = false;

                            TabCollection.TabPages.Remove(tabPage1);

                            break;
                        }

                    case "M14": //Accounting Banking Group
                        {

                            this.cbbx_busUnit.Items.AddRange(new object[] {
                        "RU0","RU1","RU4","RU5"});
                            string year, month, ActPeriod;

                            for (int m = 1; m < 24; m++)
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
                            TabCollection.TabPages.Remove(tabPage1);
                            TabCollection.TabPages.Remove(tabPage2);
                            TabCollection.TabPages.Remove(tabPage3);
                            TabCollection.TabPages.Remove(tabPage4);
                            TabCollection.TabPages.Remove(tabPage5);
                            TabCollection.TabPages.Remove(tabPage6);
                            TabCollection.TabPages.Remove(tabPage7);
                            TabCollection.TabPages.Remove(tabPage8);
                            TabCollection.TabPages.Remove(tabPage11);
                            TabCollection.TabPages.Remove(tabPage12);
                            TabCollection.TabPages.Remove(tabPage13);
                            TabCollection.TabPages.Remove(tabPage14);
                            TabCollection.TabPages.Remove(tabPage15);
                            TabCollection.TabPages.Remove(tabPage17);
                            this.pnl_accpay1.Visible = true;
                            this.pnl_accpay2.Visible = false;
                            this.pnl_casher2.Visible = true;
                            this.pnl_shift4.Visible = false;

                            break;
                        }

                    case "M15": //Accounting Banking Group (Курганская)
                        {

                            this.cbbx_busUnit.Items.AddRange(new object[] {
                        "RU0","RU1","RU4","RU5"});
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
                            TabCollection.TabPages.Remove(tabPage1);
                            TabCollection.TabPages.Remove(tabPage2);
                            TabCollection.TabPages.Remove(tabPage3);
                            TabCollection.TabPages.Remove(tabPage4);
                            TabCollection.TabPages.Remove(tabPage5);
                            TabCollection.TabPages.Remove(tabPage7);
                            TabCollection.TabPages.Remove(tabPage8);
                            TabCollection.TabPages.Remove(tabPage11);
                            TabCollection.TabPages.Remove(tabPage13);
                            TabCollection.TabPages.Remove(tabPage14);
                            TabCollection.TabPages.Remove(tabPage15);
                            TabCollection.TabPages.Remove(tabPage17);
                            this.pnl_accpay1.Visible = true;
                            this.pnl_accpay2.Visible = false;
                            this.pnl_casher2.Visible = false;
                            this.pnl_shift4.Visible = false;

                            string esender = GetSender(gv_winuser);

                            this.cbbx_Reciver.Text = esender;
                            this.cbbx_Reciver.Items.AddRange(new object[] {
                        esender});
                            break;
                        }

                    case "M16": //Payroll
                        {

                            this.cbbx_busUnit.Items.AddRange(new object[] {
                        "RU0","RU4"});
                            this.lbl_Caption.Text = "Журналы";
                            TabCollection.TabPages.Remove(tabPage1);
                            TabCollection.TabPages.Remove(tabPage2);
                            TabCollection.TabPages.Remove(tabPage3);
                            TabCollection.TabPages.Remove(tabPage4);
                            TabCollection.TabPages.Remove(tabPage5);
                            TabCollection.TabPages.Remove(tabPage6);
                            TabCollection.TabPages.Remove(tabPage7);
                            TabCollection.TabPages.Remove(tabPage8);
                            TabCollection.TabPages.Remove(tabPage9);
                            TabCollection.TabPages.Remove(tabPage10);
                            TabCollection.TabPages.Remove(tabPage11);
                            TabCollection.TabPages.Remove(tabPage13);
                            TabCollection.TabPages.Remove(tabPage14);
                            TabCollection.TabPages.Remove(tabPage15);
                            TabCollection.TabPages.Remove(tabPage17);
                            this.pnl_accpay1.Visible = false;
                            this.pnl_accpay2.Visible = false;
                            this.pnl_casher2.Visible = false;
                            this.pnl_shift4.Visible = false;

                            TabCollection.TabPages.Remove(tabPage1);

                            break;
                        }


                    case "A11": //Billing Banking group
                        {
                            this.cbbx_busUnit.Items.AddRange(new object[] {
                        "RU0"});

                            //                    tabPage2.Enabled = false;
                            this.lbl_Caption.Text = "Аллокирование";
                            TabCollection.TabPages.Remove(tabPage1);
                            TabCollection.TabPages.Remove(tabPage2);
                            TabCollection.TabPages.Remove(tabPage4);
                            TabCollection.TabPages.Remove(tabPage5);
                            TabCollection.TabPages.Remove(tabPage6);
                            TabCollection.TabPages.Remove(tabPage7);
                            TabCollection.TabPages.Remove(tabPage9);
                            TabCollection.TabPages.Remove(tabPage8);
                            TabCollection.TabPages.Remove(tabPage10);
                            TabCollection.TabPages.Remove(tabPage11);
                            TabCollection.TabPages.Remove(tabPage12);
                            TabCollection.TabPages.Remove(tabPage13);
                            TabCollection.TabPages.Remove(tabPage14);
                            TabCollection.TabPages.Remove(tabPage15); //сдвиг периода
                            TabCollection.TabPages.Remove(tabPage17);
                            this.pnl_accpay1.Visible = false;
                            this.pnl_accpay2.Visible = false;
                            this.pnl_casher2.Visible = false;
                            this.pnl_shift4.Visible = false;

                            this.btn_allocate.Enabled = false;

                            break;
                        }
                    case "U11": //Admin
                        {
                            this.cbbx_busUnit.Items.AddRange(new object[] {
                        "RU0","RU1","RU4","RU5"});


                            string year, month, ActPeriod;

                            for (int m = 1; m < 24; m++)
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
                                this.cbbx_Reciver.Items.AddRange(new object[] { emails[i][2].ToLower() });
                                // MessageBox.Show(esender);
                            }

                            this.lbl_Caption.Text = "Админ";
                            this.btn_allocate.Enabled = false;
                            this.pnl_accpay1.Visible = true;
                            this.pnl_accpay2.Visible = true;
                            this.pnl_casher2.Visible = true;
                            this.pnl_shift4.Visible = true;


                            break;
                        }
                    default: //Default
                        {
                            this.cbbx_busUnit.Items.AddRange(new object[] {
                        "NotFound"});
                            this.lbl_Caption.Text = "";
                            TabCollection.TabPages.Remove(tabPage1);
                            TabCollection.TabPages.Remove(tabPage2);
                            TabCollection.TabPages.Remove(tabPage3);
                            TabCollection.TabPages.Remove(tabPage4);
                            TabCollection.TabPages.Remove(tabPage5);
                            TabCollection.TabPages.Remove(tabPage6);
                            TabCollection.TabPages.Remove(tabPage7);
                            TabCollection.TabPages.Remove(tabPage8);
                            TabCollection.TabPages.Remove(tabPage9);
                            TabCollection.TabPages.Remove(tabPage10);
                            TabCollection.TabPages.Remove(tabPage11);
                            TabCollection.TabPages.Remove(tabPage12);
                            TabCollection.TabPages.Remove(tabPage13);
                            TabCollection.TabPages.Remove(tabPage14);
                            TabCollection.TabPages.Remove(tabPage15);
                            TabCollection.TabPages.Remove(tabPage16);
                            this.pnl_accpay1.Visible = false;
                            this.pnl_accpay2.Visible = false;
                            this.pnl_casher2.Visible = false;
                            this.pnl_shift4.Visible = false;

                            this.label1.Visible = false;
                            this.label4.Visible = false;
                            this.label8.Visible = false;

                            this.txbx_sunUser.Visible = false;
                            this.lbl_period.Visible = false;


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
            this.lbl_perShift.Text = "";
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

            int cntJournalDeleted = dal.JrnalPerDayModif(cbbx_busUnit.SelectedItem.ToString(), gv_winuser, 1);

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
                this.lbl_delval.Text = dal.JrnalPerDayModif(cbbx_busUnit.SelectedItem.ToString(), gv_winuser, 1).ToString();
            }

        }

        private void txbx_journalNumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar < 48 || e.KeyChar > 58 && e.KeyChar != 8)
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

                MainForm.SendMail("oleg.lesnitsky@dhl.com", "Инфо: " + ins_res, "Коллеги, добрый день!\n\n" + ins_res + " \n\nЭто служебное сообщение (и адрес эл.почты), просьба на него не отвечать. \nСпасибо. \n\nС Уважением, Олег Лесницкий.\nmailto:oleg.lesnitsky@dhl.com", this.cbbx_busUnit.ToString(), 0);


                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + ins_res + "\n" + rtxbx_info.Text;

                btn_addCourier.Enabled = true;
                this.lbl_adedCourier.Text = dal.CourierPerDayIns(cbbx_busUnit.SelectedItem.ToString(), gv_winuser).ToString();

                this.txbx_courierLookup.Text = txbx_fiocourier.Text.ToString().Substring(0, 15).ToUpper();

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

            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

            Excel.Application xlApp = new Excel.Application();
            string[] FileNames = new[] { "" };

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
                MessageBox.Show("Из папки <" + this.txbx_pathForList.Text + "> в Excel выгружено " + (FileNames.Length - 1) + " имен(и) файлов. ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                xlApp.Visible = true;

                xlApp.Interactive = true;
                xlApp.ScreenUpdating = true;
                xlApp.UserControl = true;
            }

        }


        //Вывод списка Каталогов
        private void btn_upNDirToExcel_Click(object sender, EventArgs e)
        {

            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

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
                this.rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Из папки <" + this.txbx_pathForList.Text + "> в Excel выгружено " + (DirNames.Length - 1) + " имен(и) каталогов. " + "\n" + rtxbx_info.Text;
                MessageBox.Show("Из папки <" + this.txbx_pathForList.Text + "> в Excel выгружено " + (DirNames.Length - 1) + " имен(и) Каталогов. ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

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

        private string[] GetFiles3(string path)
        {
            FileInfo[] fi;
            ArrayList al = new ArrayList();

            DirectoryInfo di = new DirectoryInfo(path);
            fi = di.GetFiles("*.doc");
            foreach (FileInfo f in fi)
            {
                al.Add(f.Name);
            }
            string[] files = (string[])al.ToArray(typeof(string));
            return files;
        }

        private string[] GetFiles4(string agnCode, string path)
        {
            FileInfo[] fi;
            ArrayList al = new ArrayList();

            DirectoryInfo di = new DirectoryInfo(path);
            fi = di.GetFiles(agnCode+"*.pdf");
            foreach (FileInfo f in fi)
            {
                al.Add(path+"\\"+f.Name);
            }
            string[] files = (string[])al.ToArray(typeof(string));
            return files;
        }

        private string[] GetFiles4inst(string agnCode, string path)
        {
            FileInfo[] fi;
            ArrayList al = new ArrayList();

            DirectoryInfo di = new DirectoryInfo(path);
            fi = di.GetFiles(agnCode + "*.docx");
            foreach (FileInfo f in fi)
            {
                al.Add(path + "\\" + f.Name);
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

        static public void SendMail(string reciever, string mesSubj, string mesBody, string busUnit, int additionalMail)
        {
            //smtp сервер
            string smtpHost = "gateway.dhl.com";
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

            //!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            /*
            if (busUnit == "M11")
            subject = "Инфо: ПЕРИОД в 1C Trade AO ";
            else if (busUnit == "E11")
            subject = "Инфо: ПЕРИОД в 1C Trade OOO ";

            if (OpenPer)
                subject = subject + " Открыт[+]. ";
            else
                subject = subject + " Закрыт[-]. ";
            string subject = mesSubj;
            */
            //Текст письма
            //string body = "Привет! \n\n\n Это тестовое письмо от C Sharp";


            //Создаем сообщение
            MailMessage mess = new MailMessage(from, reciever, mesSubj, mesBody);
            mess.BodyEncoding = Encoding.UTF8;
            mess.CC.Add("oleg.lesnitsky@dhl.com");
            //получатели
            mess.CC.Add("albina.kutlueva@dhl.com");

            if (additionalMail == 1)
                mess.CC.Add("oleg.lesnitsky@dhl.com");

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

        static public string SendMail3(string[] reciever, string mesSubj, string mesBody, string[] fColl)
        {
            //smtp сервер
            string smtpHost = "smtp.prg-dc.dhl.com";
            //smtp порт
            int smtpPort = 25;
            //логин
            string login = "ruhrpadm";
            //пароль
            string pass = "Rh121212";

            string result;

            //создаем подключение
            SmtpClient client = new SmtpClient(smtpHost, smtpPort);
            client.Credentials = new NetworkCredential(login, pass);

            //От кого письмо
            string from = "ruhrpadm@dhl.com";
            //Кому письмо
            //string to = "oleg.lesnitsky@dhl.com";
            //Тема письма

            //!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            /*
            if (busUnit == "M11")
            subject = "Инфо: ПЕРИОД в 1C Trade AO ";
            else if (busUnit == "E11")
            subject = "Инфо: ПЕРИОД в 1C Trade OOO ";

            if (OpenPer)
                subject = subject + " Открыт[+]. ";
            else
                subject = subject + " Закрыт[-]. ";
            string subject = mesSubj;
            */
            //Текст письма
            //string body = "Привет! \n\n\n Это тестовое письмо от C Sharp";


            //Создаем сообщение
            MailMessage mess = new MailMessage(from, reciever[0], mesSubj, mesBody);

            foreach (string eml in reciever)
            {
                if (eml != reciever[0])
                    mess.CC.Add(eml);
            }

            mess.BodyEncoding = Encoding.UTF8;
            mess.CC.Add("oleg.lesnitsky@dhl.com");
            //
            mess.CC.Add("albina.kutlueva@dhl.com");

            foreach (string fName in fColl)
            {
                Attachment att = new Attachment(fName);
                mess.Attachments.Add(att);
            }
            //            MailMessage mess2 = new MailMessage(from, "oleg.lesnitsky@dhl.com", subject, mesBody);
            //            mess2.BodyEncoding = Encoding.UTF8;
            try
            {
                client.Send(mess);
                result = " '" + fColl.Length.ToString() + "' файлов PDF отправлены для <" + reciever + ">";
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.ReadKey();
                result = " ошибка отправки файлов <" + reciever + ">;";
                return result;
            }

        }


        static public string SendMail4(string[] reciever, string mesSubj, string mesBody, string[] fColl, string[] insrColl)
        {
            //smtp сервер
            string smtpHost = "smtp.prg-dc.dhl.com";
            //smtp порт
            int smtpPort = 25;
            //логин
            string login = "ruhrpadm";
            //пароль
            string pass = "Rh121212";

            string result;

            //создаем подключение
            SmtpClient client = new SmtpClient(smtpHost, smtpPort);
            client.Credentials = new NetworkCredential(login, pass);

            //От кого письмо
            string from = "ruhrpadm@dhl.com";
            //Кому письмо
            //string to = "oleg.lesnitsky@dhl.com";
            //Тема письма

            //!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            /*
            if (busUnit == "M11")
            subject = "Инфо: ПЕРИОД в 1C Trade AO ";
            else if (busUnit == "E11")
            subject = "Инфо: ПЕРИОД в 1C Trade OOO ";

            if (OpenPer)
                subject = subject + " Открыт[+]. ";
            else
                subject = subject + " Закрыт[-]. ";
            string subject = mesSubj;
            */
            //Текст письма
            //string body = "Привет! \n\n\n Это тестовое письмо от C Sharp";


            //Создаем сообщение
            MailMessage mess = new MailMessage(from, reciever[0], mesSubj, mesBody);

            foreach (string eml in reciever)
            {
                if (eml != reciever[0])
                    mess.CC.Add(eml);
            }

            mess.BodyEncoding = Encoding.UTF8;
            mess.CC.Add("oleg.lesnitsky@dhl.com");
            mess.CC.Add("albina.kutlueva@dhl.com");

            foreach (string fName in insrColl)
            {
                Attachment att = new Attachment(fName);
                mess.Attachments.Add(att);
            }

            foreach (string fName in fColl)
            {
                Attachment att = new Attachment(fName);
                mess.Attachments.Add(att);
            }
            //            MailMessage mess2 = new MailMessage(from, "oleg.lesnitsky@dhl.com", subject, mesBody);
            //            mess2.BodyEncoding = Encoding.UTF8;
            try
            {
                client.Send(mess);
                result = " '" + fColl.Length.ToString() + "' файлов PDF отправлены для <" + reciever + ">";
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.ReadKey();
                result = " ошибка отправки файлов <" + reciever + ">;";
                return result;
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
                MessageBox.Show(" Укажите значения в поле БизнесЮнит [RU0/RU1 - Trade АО или RU4/RU5 - Trade ООО]:", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (busUnit.ToString() != "RU1" && busUnit.ToString() != "RU5" && busUnit.ToString() != "RU0" && busUnit.ToString() != "RU4")
            {
                MessageBox.Show(" Укажите значения в поле БизнесЮнит [RU0/RU1 - Trade АО или RU4/RU5 - Trade ООО]:", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
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

            string subject = "";

            if (busUnit == "RU0")
                subject = "Инфо: ПЕРИОД в 1C Trade AO Открыт[+]. ";
            else if (busUnit == "RU4")
                subject = "Инфо: ПЕРИОД в 1C Trade OOO Открыт[+]. ";

            //Подключиться к базе и поменять период по параметрам (период, регион)
            //Записать в лог (кто, во-сколько, период, регион, причина)
            //Отправить оповещение заинтересованным пользователям
            SendMail(this.cbbx_Reciver.Text.ToString(),
                    subject,
                    "Коллеги, добрый день!\n\nПериод: " + this.cbbx_TradePeriod.Text.ToString() + " \nдля региона: " + this.cbbx_region.SelectedItem.ToString() + " \nпо причине: <" + this.txbx_ComntPeriod.Text.ToString() + "> \nоткрыт.\nНе забудьте Закрыть период!\n\nЭто служебное сообщение (и адрес эл.почты), просьба на него не отвечать. \nСпасибо. \n\nС Уважением, Олег Лесницкий.\nmailto:oleg.lesnitsky@dhl.com",
                    this.cbbx_busUnit.SelectedItem.ToString(),0);

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
                MessageBox.Show(" Укажите значения в поле БизнесЮнит  [RU0/RU1 - Trade АО или RU4/RU5 - Trade ООО]:", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (busUnit.ToString() != "RU1" && busUnit.ToString() != "RU5" && busUnit.ToString() != "RU0" && busUnit.ToString() != "RU4")
            {
                MessageBox.Show(" Укажите значения в поле БизнесЮнит [ [RU0/RU1 - Trade АО или RU4/RU5 - Trade ООО]:", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
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

            string subject = "";

            if (busUnit == "RU1")
                subject = "Инфо: ПЕРИОД в 1C Trade AO Закрыт[+]. ";
            else if (busUnit == "RU5")
                subject = "Инфо: ПЕРИОД в 1C Trade OOO Закрыт[+]. ";
            else if (busUnit == "RU0")
                subject = "Инфо: ПЕРИОД в 1C Trade AO Закрыт[+]. ";
            else if (busUnit == "RU4")
                subject = "Инфо: ПЕРИОД в 1C Trade OOO Закрыт[+]. ";

            SendMail(this.cbbx_Reciver.Text.ToString(),
                subject,
                "Коллеги, добрый день!\n\nПериод для региона: " + this.cbbx_region.SelectedItem.ToString() + "\nзакрыт.\n\nЭто служебное сообщение (и адрес эл.почты), просьба на него не отвечать.\n\nС Уважением, Олег Лесницкий.\nmailto:oleg.lesnitsky@dhl.com",
                this.cbbx_busUnit.SelectedItem.ToString(),0);

        }

        private void btn_rename_Click(object sender, EventArgs e)
        {
            string[] FileNames = new[] { "" };

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
                    fl.MoveTo(this.txbx_pathForList.Text + "\\" + fl.Name.Substring(0, 9) + ".tif");

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


        private void button1_Click_1(object sender, EventArgs e)
        {
            int Searched = 0;
            string searchFoldName = "380002669";

            string[] FoldNames = new[] { "" };
            DirectoryInfo df = new DirectoryInfo(this.txbx_path.Text);

            FoldNames = GetDirs(this.txbx_path.Text);

            //C:\OL2\dpo_management.xls
            for (int FoldInd = 0; FoldInd < FoldNames.Length; FoldInd++)
            {

                if (FoldNames[FoldInd] == searchFoldName)
                {
                    Searched = Searched + 1;
                }



                //fl = new FileInfo(FileNames[fileInd]);
                //fl.MoveTo(this.txbx_pathForList.Text + "\\" + fl.Name.Substring(0, 8));
            }

            MessageBox.Show(Searched.ToString());

        }

        private void cbtn_convert_Click(object sender, EventArgs e)
        {
            string[] FileNames = new[] { "" };
            Excel.Workbook ObjWorkBook = null;
            Excel.Worksheet ObjWorkSheet = null;
            Excel.Range ObjWorkRange = null;

            DirectoryInfo df = new DirectoryInfo(this.txbx_pathForListSF.Text + "\\xls");
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

            for (int i = 0; i < users.Length - 2; i = i + 3)
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
            if (busUnit == null || busUnit.ToString() != "RU0")
            {
                MessageBox.Show(" Укажите значения в поле БизнесЮнит [RU0 - Trade АО]:", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            DateTime rate_datetime;
            rate_datetime = DateTime.Today.AddDays(1);


            if (this.dtpk_dateRate.Text != null)
                rate_datetime = DateTime.Parse(this.dtpk_dateRate.Text.ToString(), CultureInfo.CreateSpecificCulture("ru-RU"));


            if (rate_datetime > DateTime.Today.AddDays(10) || rate_datetime < DateTime.Today.AddDays(-31))
            {
                MessageBox.Show(" Значение Даты Курсов должно быть \n не позднее текущей даты и не более 31 дня ранее...\n Проверьте значение даты!", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            this.rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + dal.LoadRateToSun(rate_datetime, gv_winuser, this.txbx_sunUser.Text) + "\n" + rtxbx_info.Text;

        }

        private void txbx_journaUselNumber_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar < 48 || e.KeyChar > 58 && e.KeyChar != 8)
                e.Handled = true;
        }

        private void txbx_journaUselNumber_TextChanged(object sender, EventArgs e)
        {
            if (this.txbx_journaUselNumber.Text.Length > 4)
                this.btn_offUse.Enabled = true;
        }

        private void btn_offUse_Click(object sender, EventArgs e)
        {

            object busUnit = cbbx_busUnit.SelectedItem;
            if (busUnit == null)
            {
                MessageBox.Show(" Укажите значения в поле БизнесЮнит: ", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            btn_offUse.Enabled = false;
            this.rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + dal.JournalInUseOff(cbbx_busUnit.SelectedItem.ToString(), this.txbx_journaUselNumber.Text.ToString(), lbl_period.Text, txbx_sunUser.Text, lbl_role.Text, gv_winuser) + "\n" + rtxbx_info.Text;
            btn_offUse.Enabled = true;
            //this.lbl_delval.Text = dal.JrnalPerDayModif(cbbx_busUnit.SelectedItem.ToString(), gv_winuser).ToString();

        }

        private void btn_JournalWork_Click(object sender, EventArgs e)
        {
            object busUnit = cbbx_busUnit.SelectedItem;
            if (busUnit == null)
            {
                MessageBox.Show(" Укажите значения в поле БизнесЮнит [RU0 - Trade АО]:", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            object jrnTemp = cmbx_journTempl.SelectedItem;
            if (jrnTemp == null || jrnTemp.ToString().Length <= 6)
            {
                MessageBox.Show(" Укажите значения в поле Шаблон просмотра ", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            jrnTemp = this.txbx_reference2;
            if (jrnTemp == null || jrnTemp.ToString().Length == 0)
            {
                this.txbx_reference2.Text = "%";
            }

            frm_oprJournal orpJournal = new frm_oprJournal(this.txbx_journalNumber3.Text, this.txbx_reference2.Text, this.cmbx_journTempl.SelectedIndex + 1, this.cmbx_journTempl.SelectedItem.ToString(), busUnit.ToString());
            orpJournal.Show();
        }

        private void txbx_journalNumber3_TextChanged(object sender, EventArgs e)
        {
            //if (this.txbx_journalNumber3.Text.Length > 3 && this.txbx_reference2.Text.Length > 2)
            if (this.txbx_journalNumber3.Text.Length > 3)
                this.btn_JournalWork.Enabled = true;
        }

        private void txbx_reference2_TextChanged(object sender, EventArgs e)
        {
            //if (this.txbx_journalNumber3.Text.Length > 3 && this.txbx_reference2.Text.Length > 2)            
            if (this.txbx_reference2.Text.Length > 2)
                this.btn_JournalWork.Enabled = true;
        }

        private void btn_supUnblock_Click(object sender, EventArgs e)
        {
            object busUnit = cbbx_busUnit.SelectedItem;
            if (busUnit == null)
            {
                MessageBox.Show(" Укажите значения в поле БизнесЮнит [RU1, RU0 - АО или RU5 - ООО]:", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (this.txbx_supCode2.Text.ToString().Trim().Length < 8)
            {
                MessageBox.Show(" Укажите корректное значение в поле Код Поставщика ", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                //MessageBox.Show(this.txbx_supCode2.ToString().Length.ToString(),"1212");
                return;
            }

            btn_supUnblock.Enabled = false;
            btn_supUnblockTrade.Enabled = false;

            this.rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + dal.supUnblock(cbbx_busUnit.SelectedItem.ToString(), this.txbx_supCode2.Text.ToString(), txbx_sunUser.Text, lbl_role.Text, gv_winuser) + "\n" + rtxbx_info.Text;
            btn_supUnblock.Enabled = true;
            //this.lbl_delval.Text = dal.JrnalPerDayModif(cbbx_busUnit.SelectedItem.ToString(), gv_winuser).ToString();

        }

        private void btn_supUnblockTrade_Click(object sender, EventArgs e)
        {
            object busUnit = cbbx_busUnit.SelectedItem;
            if (busUnit == null || (busUnit.ToString() != "RU0" && busUnit.ToString() != "RU4"))
            {
                MessageBox.Show(" Укажите значения в поле БизнесЮнит [RU0 = Trade АО] или [RU4 = Trade АО]", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }


            if (this.txbx_supCode2.Text.ToString().Trim().Length < 8)
            {
                MessageBox.Show(" Укажите корректное значение в поле Код Поставщика ", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (this.txbx_reasonUnbl.Text.ToString().Trim().Length < 10)
            {
                MessageBox.Show(" Укажите Основание разблокировки ", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }


            btn_supUnblock.Enabled = false;
            btn_supUnblockTrade.Enabled = false;
            this.rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + dal.supUnblockTrade(cbbx_busUnit.SelectedItem.ToString(), this.txbx_supCode2.Text.ToString(), txbx_sunUser.Text, lbl_role.Text, gv_winuser, txbx_reasonUnbl.Text.ToString().Trim()) + "\n" + rtxbx_info.Text;
            //btn_supUnblock.Enabled = true;
            btn_supUnblockTrade.Enabled = true;

        }

        private void cbtn_jnumclear_Click(object sender, EventArgs e)
        {
            txbx_journalNumber3.Text = "";
            txbx_reference2.Text = "";
            txbx_journalNumber3.Focus();

        }

        private void btn_getDname_Click(object sender, EventArgs e)
        {
            if (txbt_ipAdr.Text.ToString().Trim().Length >= 10)
            {
                this.txbt_domainName.Text = DoGetHostEntry(this.txbt_ipAdr.Text.ToString());
            }
            else
            {
                MessageBox.Show(" Укажите корректное значение в поле IP ", "  SUNPLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
        }

        public static string DoGetHostEntry(string hostIp)
        {
            IPAddress addr = IPAddress.Parse(hostIp);
            IPHostEntry EntryName = Dns.GetHostEntry(addr);

            string hname = EntryName.HostName.ToString();
            return hname.Remove(15);
        }

        private void cbtn_copyBuf_Click(object sender, EventArgs e)
        {
            if (this.txbt_domainName.Text.ToString().Trim().Length >= 15)
            {
                this.txbt_domainName.Text = DoGetHostEntry(this.txbt_ipAdr.Text.ToString());
            }

        }

        private void txbx_journalNumber3_MouseClick(object sender, MouseEventArgs e)
        {
            DigitConsole dc = new DigitConsole();
            dc.Show();
        }

        private void btn_getVUsers_Click(object sender, EventArgs e)
        {
            this.rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + dal.VisionUserManage(1, "M11", gv_winuser) + "\n" + rtxbx_info.Text;
        }

        private void btn_clearVUsrs_Click(object sender, EventArgs e)
        {
            this.rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + dal.VisionUserManage(3, "M11", gv_winuser) + "\n" + rtxbx_info.Text;
        }

        private void txbx_journalNumber4_TextChanged(object sender, EventArgs e)
        {
            if (this.txbx_journalNumber4.Text.Length > 2)
                this.btn_prdCng.Enabled = true;
        }

        private void btn_prdCng_Click(object sender, EventArgs e)
        {
            if (txbx_sunUser.Text == "не опред." || lbl_period.Text == "не опред.")
            {
                MessageBox.Show(" Проверьте наличие значения в поле Бизнес Юнит! ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (txbx_journalNumber4.Text.ToString() == string.Empty || txbx_journalNumber4.Text.ToString() == "")
            {
                MessageBox.Show(" Введите номер журнала! ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (txbx_newPeriod.Text.ToString() == string.Empty || txbx_newPeriod.Text.ToString() == "")
            {
                MessageBox.Show(" Введите поле Новый период! ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (txbx_newPeriod.Text.ToString() == string.Empty || txbx_newPeriod.Text.ToString() == "")
            {
                MessageBox.Show(" Введите поле Новый период! ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (txbx_reasonMov.Text.ToString() == string.Empty || txbx_reasonMov.Text.ToString() == "")
            {
                MessageBox.Show(" Введите поле Основание изменения периода! ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            int additionalFlagSend = 0;

            if (chbx_mCpa.Checked)
            {
                additionalFlagSend = 1;
            }



            int cntJournalShift = dal.JrnalPerDayModif(cbbx_busUnit.SelectedItem.ToString(), gv_winuser, 2);

            if (cntJournalShift == 3 && gv_role != "R01" && gv_role != "R05")
            //            if (int.Parse(lbl_limit.Text) == cntJournalDeleted)
            {
                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Сдвиг журнала для " + cbbx_busUnit.SelectedItem.ToString() + " невозможен, превышение лимита на СдвигПериода за текущий день, либо неверный БизнесЮнит (можно указать только RU0,RU4);\n" + rtxbx_info.Text;
            }
            else
            {
                btn_prdCng.Enabled = false;
                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] " + dal.journalShiftLayout(cbbx_busUnit.SelectedItem.ToString(), txbx_journalNumber4.Text.ToString(), lbl_period.Text, txbx_sunUser.Text, lbl_role.Text, gv_winuser, txbx_newPeriod.Text, txbx_reasonMov.Text, additionalFlagSend) + "\n" + rtxbx_info.Text;
                btn_prdCng.Enabled = true;
                this.lbl_perShift.Text = dal.JrnalPerDayModif(cbbx_busUnit.SelectedItem.ToString(), gv_winuser, 2).ToString();
            }


        }

        private void txbx_journalNumber4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar < 48 || e.KeyChar > 58 && e.KeyChar != 8)
                e.Handled = true;
        }

        private void btn_getXMLinfo_Click(object sender, EventArgs e)
        {

            // btn_getXMLinfo_Click (в string получаем протокол ошибок)
            // --> Workxml.GetXMLFilesInfo
            // ------> CheckXMLprotocol, CheckJrnalSrc
            // ----------> getSunXMLErrors



            //rtxbx_info.Text = Workxml.GetXMLFilesInfo(@"\\RUMOWWSX12031\transferlogs", "EVA") + rtxbx_info.Text;

            //rtxbx_info.Text = Workxml.GetXMLFilesInfo(@"C:\OL2\C#\SUN_xml\", "EVA") + rtxbx_info.Text;


            //dgv_XMLProtocol.GridColor = Color.Black;
            //dgv_XMLProtocol.DataSource = Workxml.GetXMLDataTableProtocol(@"\\RUMOWWSX12031\transferlogs", "ASO");


            //DataTable Table = Workxml.GetXMLDataTableProtocol(@"\\RUMOWWSX12031\transferlogs", "EKZ");

            //

            if (txbx_sunUser.Text == "не опред." || lbl_period.Text == "не опред.")
            {
                MessageBox.Show(" Проверьте наличие значения в поле Бизнес Юнит! ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            if (txbx_jrnalSrce2.Text.Length < 3)
            {
                MessageBox.Show(" Укажите Journal Source!", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            Workxml Workxml = new Workxml();
            DAL DAL = new DAL();

            //MessageBox.Show("01 Запуск процедуры1");

            //DataTable Table = Workxml.GetXMLDataTableProtocol(@"\\RUMOWWSX12031\transferlogs", txbx_sunUser.Text.Trim());
            DataTable Table = Workxml.GetXMLDataTableProtocol(@"\\RUMOWWSX12031\transferlogs", txbx_jrnalSrce2.Text.ToUpper());

            frm_SunXMLprotocol XMLprotocol = new frm_SunXMLprotocol(ref Table);

            //
            XMLprotocol.Show();

            if (Table.Rows.Count > 0)
                DAL.RecordLog("SUNPLUS_06", "SUNTDXML", "winuser", "000000", txbx_sunUser.Text, this.cbbx_busUnit.SelectedItem.ToString(), 0, 0, 0, "PROT_without_ROWS", "ПротоколОшибок");
            else
                DAL.RecordLog("SUNPLUS_06", "SUNTDXML", "winuser", "000000", txbx_sunUser.Text, this.cbbx_busUnit.SelectedItem.ToString(), 0, 0, 0, "PROT_with_ROWS", "ПротоколОшибок");

            rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Запрос к Протоколам загружаемых файлов отработал.\n" + rtxbx_info.Text;

        }

        private void cbtn_convertPDF_Click(object sender, EventArgs e)
        {

            object oMissing = System.Reflection.Missing.Value;


            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

            try
            {

                // Get list of Word files in specified directory
                DirectoryInfo dirInfo = new DirectoryInfo(this.txbx_pathForListDOCfiles.Text);
                FileInfo[] wordFiles = dirInfo.GetFiles("*.doc");

                word.Visible = false;
                word.ScreenUpdating = false;

                int filesCount;
                filesCount = 0;

                if (wordFiles.Length == 0)
                {

                    rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Файлы DOC в папке <" + this.txbx_pathForListDOCfiles.Text + "> не обнаружены.\n" + rtxbx_info.Text;
                    MessageBox.Show("Файлы DOC в папке <" + this.txbx_pathForListDOCfiles.Text + "> не обнаружены.", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;

                }
                
                foreach (FileInfo wordFile in wordFiles)
                {
                    // Cast as Object for word Open method
                    Object filename = (Object)wordFile.FullName;

                    // Use the dummy value as a placeholder for optional arguments
                    Word.Document doc = word.Documents.Open(ref filename, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                    doc.Activate();

                    object outputFileName = wordFile.FullName.Replace(".doc", ".pdf");
                    object fileFormat = Word.WdSaveFormat.wdFormatPDF;

                    // Save document into PDF Format
                    doc.SaveAs(ref outputFileName,
                        ref fileFormat, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                    // Close the Word document, but leave the Word application open.
                    // doc has to be cast to type _Document so that it will find the
                    // correct Close method.                
                    object saveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;
                    ((Word._Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
                    doc = null;

                    filesCount = filesCount + 1;
                }

                // word has to be cast to type _Application so that it will find
                // the correct Quit method.
                ((Word._Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
                word = null;

                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "]  " + filesCount.ToString() + " DOC файлов обработано\n" + rtxbx_info.Text;

                MessageBox.Show("Конвертация DOC->PDF: " + filesCount.ToString() + " DOC файлов обработано!", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "]  " + "Проверьте наличие каталога " + this.txbx_pathForListDOCfiles.Text + " \n" + rtxbx_info.Text;

                MessageBox.Show("Конвертация DOC->PDF:\n" + "Проверьте наличие каталога " + this.txbx_pathForListDOCfiles.Text + "\nОбработка остановлена", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);


            }

        }

        private void btn_sendPdf_Click(object sender, EventArgs e)
        {

                DirectoryInfo dirInfo = new DirectoryInfo(this.txbx_pathForListDOCfiles.Text);
                FileInfo[] pdfFiles = dirInfo.GetFiles("*.pdf");

                if (pdfFiles.Length == 0)
                {

                    rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Файлы PDF в папке <" + this.txbx_pathForListDOCfiles.Text + "> не обнаружены.\n" + rtxbx_info.Text;
                    MessageBox.Show("Файлы PDF в папке <" + this.txbx_pathForListDOCfiles.Text + "> не обнаружены.", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;

                }


                string[] agnFileColl;
                string[] instrFileColl;

                string[] agnMail;
                string strAgnts = "";
                string agnMessBody;
                string strRepPeriod;

                strRepPeriod = "02.2019";

                  //                agnMessBody = "Здравствуйте,\n\nTEST TEST TEST\n\nВо вложении документы PDF для Электронного Обмена за 10.2018 период.\nПросьба сконтактировать с Представителем АО ''ДХЛ Интернешнл'' для тестовой отправки данных документов через Вашу систему Электронного Документооборота. \nЭто служебное сообщение (и адрес эл.почты), просьба на него не отвечать.\n\nС Уважением, DHL Express.";
                  //                agnMessBody = "Здравствуйте,\n\nTEST TEST TEST\n\nВ связи с введением электронного документооборота между нашими организациями прошу вас принять участие в тестовой отправке документов через Вашу систему Электронного Документооборота.\n\nВо вложении документы PDF (акты) для Электронного Обмена за 10.2018 период. \nА также инструкция по их внесению в систему Электронного документооборота.\n\nЭто служебное сообщение (и адрес эл.почты), просьба на него не отвечать.\n\nПо вопросам можно побращаться Phone: +7 495 9561001 :\nАльбина Кутлуева mailto:albina.kutlueva@dhl.com [Добавочный: 3000] или\nОлег Лесницкий mailto:oleg.lesnitsky@dhl.com [Добавочный: 2416]\n\nС Уважением, DHL Express.";
                  agnMessBody = "Здравствуйте,\n\nВо вложении документы PDF (акты) для Электронного Обмена за " + strRepPeriod + " период. \n\nЭто служебное сообщение (и адрес эл.почты), просьба на него не отвечать.\n\nПо вопросам можно побращаться Phone: +7 495 9561001 :\nАльбина Кутлуева mailto:albina.kutlueva@dhl.com [Добавочный: 3000] или\nОлег Лесницкий mailto:oleg.lesnitsky@dhl.com [Добавочный: 2416]\n\nС Уважением, DHL Express.";

            //                string[] AgnCodes = new[] { "LNX", "RZN", "URS", "VGD", "TBW" };

            if (this.txbx_listAgnts.Text.Length >= 3)
                strAgnts = this.txbx_listAgnts.Text.Trim();

            // string[] AgnCodes = new[] {"JOK", "CSY"};
            string[] AgnCodes = strAgnts.Split(' ');
            //                  string[] AgnCodes = new[] { "CEK"};

            //                string[] AgnCodes = new[] { "LNX", "RZN", "URS", "VGD", "TBW", "CEK" };

            //string[] AgnCodes = new[] { "JOK", "CSY", "LNX", "RZN", "URS", "VGD", "TBW" , "LPK", "MQF", "AER"};



            instrFileColl = GetFiles4inst("Инструкция", this.txbx_pathForListDOCfiles.Text);

                foreach (string AgnCod in AgnCodes)
                {
                    switch (AgnCod)
                    {

                        case "CEK":
                            {
                                agnFileColl = GetFiles4(AgnCod, this.txbx_pathForListDOCfiles.Text);

                                if (agnFileColl.Length != 0)
                                {
                                    //agnMail = "oleg.lesnitsky@dhl.com";
                                    //
                                    agnMail = new[] { "serg@academ-dhl.com.ru", "Alexey.Gomzyakov@dhl.ru" };

                                    rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Для " + AgnCod + SendMail4(agnMail, "Агентские документы PDF  [Агент: " + AgnCod + "] период " + strRepPeriod, agnMessBody, agnFileColl, instrFileColl) + " из каталога <" + this.txbx_pathForListDOCfiles.Text + "> \n" + rtxbx_info.Text;
                                }

                                break;
                            }

                        case "OZR":
                            {
                                agnFileColl = GetFiles4(AgnCod, this.txbx_pathForListDOCfiles.Text);

                                if (agnFileColl.Length != 0)
                                {
                                    //agnMail = "oleg.lesnitsky@dhl.com";
                                    //
                                    agnMail = new[] { "presentozr@mail.ru", "ruozkagt@dhl.com", "Alexey.Gomzyakov@dhl.ru" };

                                    rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Для " + AgnCod + SendMail4(agnMail, "Агентские документы PDF  [Агент: " + AgnCod + "] период " + strRepPeriod, agnMessBody, agnFileColl, instrFileColl) + " из каталога <" + this.txbx_pathForListDOCfiles.Text + "> \n" + rtxbx_info.Text;
                            }

                                break;
                            }


                        case "TVE":
                            {
                                agnFileColl = GetFiles4(AgnCod, this.txbx_pathForListDOCfiles.Text);

                                if (agnFileColl.Length != 0)
                                {
                                    //agnMail = "oleg.lesnitsky@dhl.com";
                                    //
                                    agnMail = new[] { "Marina.Bibikova@agent-dhl.com", "buh.dhl-tver@inbox.ru", "Alexandr.Pavlov@dhl.com" };

                                    rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Для " + AgnCod + SendMail4(agnMail, "Агентские документы PDF  [Агент: " + AgnCod + "] период " + strRepPeriod, agnMessBody, agnFileColl, instrFileColl) + " из каталога <" + this.txbx_pathForListDOCfiles.Text + "> \n" + rtxbx_info.Text;


                            }

                                break;
                            }


                        case "LNX":
                            {
                                agnFileColl = GetFiles4(AgnCod, this.txbx_pathForListDOCfiles.Text);

                                if (agnFileColl.Length != 0)
                                {
                                    //agnMail = "oleg.lesnitsky@dhl.com";
                                    //
                                    agnMail = new[] { "stavrova55@gmail.com", "rulnxagt@dhl.com", "Sergey.Kurlanov@dhl.ru" };

                                    rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Для " + AgnCod + SendMail4(agnMail, "Агентские документы PDF  [Агент: " + AgnCod + "] период " + strRepPeriod, agnMessBody, agnFileColl, instrFileColl) + " из каталога <" + this.txbx_pathForListDOCfiles.Text + "> \n" + rtxbx_info.Text;


                            }

                                break;
                            }

                        case "RZN":
                            {
                                agnFileColl = GetFiles4(AgnCod, this.txbx_pathForListDOCfiles.Text);

                                if (agnFileColl.Length != 0)
                                {
                                    //agnMail = "oleg.lesnitsky@dhl.com";
                                    //
                                    agnMail = new[] { "rurznagt@dhl.com", "vasilievalilya@mail.ru", "Sergey.Kurlanov@dhl.ru" };

                                    rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Для " + AgnCod + SendMail4(agnMail, "Агентские документы PDF  [Агент: " + AgnCod + "] период " + strRepPeriod, agnMessBody, agnFileColl, instrFileColl) + " из каталога <" + this.txbx_pathForListDOCfiles.Text + "> \n" + rtxbx_info.Text;
                            }

                                break;
                            }


                        case "URS":
                            {
                                agnFileColl = GetFiles4(AgnCod, this.txbx_pathForListDOCfiles.Text);

                                if (agnFileColl.Length != 0)
                                {
                                    //agnMail = "oleg.lesnitsky@dhl.com";
                                    //
                                    agnMail = new[] { "dhl.kursk@mail.ru", "ruursagt@dhl.com", "Sergey.Kurlanov@dhl.ru" };

                                                                        rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Для " + AgnCod + SendMail4(agnMail, "Агентские документы PDF  [Агент: " + AgnCod + "] период " + strRepPeriod, agnMessBody, agnFileColl, instrFileColl) + " из каталога <" + this.txbx_pathForListDOCfiles.Text + "> \n" + rtxbx_info.Text;
                                }

                                break;
                            }

                        case "VGD":
                            {
                                agnFileColl = GetFiles4(AgnCod, this.txbx_pathForListDOCfiles.Text);

                                if (agnFileColl.Length != 0)
                                {
                                    //agnMail = "oleg.lesnitsky@dhl.com";
                                    //
                                    agnMail = new[] { "ruvgdagt@dhl.com", "Timur.Bildanov@dhl.com" };


                                                                        rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Для " + AgnCod + SendMail4(agnMail, "Агентские документы PDF  [Агент: " + AgnCod + "] период " + strRepPeriod, agnMessBody, agnFileColl, instrFileColl) + " из каталога <" + this.txbx_pathForListDOCfiles.Text + "> \n" + rtxbx_info.Text;
                                }

                                break;
                            }

                        case "TBW":
                            {
                                agnFileColl = GetFiles4(AgnCod, this.txbx_pathForListDOCfiles.Text);

                                if (agnFileColl.Length != 0)
                                {
                                    //agnMail = "oleg.lesnitsky@dhl.com";
                                    //
                                    agnMail = new[] { "rutbwagt@dhl.com", "Sergey.Kurlanov@dhl.ru" };

                                                                        rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Для " + AgnCod + SendMail4(agnMail, "Агентские документы PDF  [Агент: " + AgnCod + "] период " + strRepPeriod, agnMessBody, agnFileColl, instrFileColl) + " из каталога <" + this.txbx_pathForListDOCfiles.Text + "> \n" + rtxbx_info.Text;
                                }

                                break;
                            }


                        case "LPK":
                            {
                                agnFileColl = GetFiles4(AgnCod, this.txbx_pathForListDOCfiles.Text);

                                if (agnFileColl.Length != 0)
                                {
                                    //agnMail = "oleg.lesnitsky@dhl.com";
                                    //
                                    agnMail = new[] { "rulpkagt@dhl.com", "Alexandr.Pavlov@dhl.com" };

                                                                        rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Для " + AgnCod + SendMail4(agnMail, "Агентские документы PDF  [Агент: " + AgnCod + "] период " + strRepPeriod, agnMessBody, agnFileColl, instrFileColl) + " из каталога <" + this.txbx_pathForListDOCfiles.Text + "> \n" + rtxbx_info.Text;
                                }

                                break;
                            }


                    case "MQF":
                        {
                            agnFileColl = GetFiles4(AgnCod, this.txbx_pathForListDOCfiles.Text);

                            if (agnFileColl.Length != 0)
                            {
                                //agnMail = "oleg.lesnitsky@dhl.com";
                                //
                                agnMail = new[] { "dhlmgn@gmail.com", "Alexey.Gomzyakov@dhl.ru" };

                                                                    rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Для " + AgnCod + SendMail4(agnMail, "Агентские документы PDF  [Агент: " + AgnCod + "] период " + strRepPeriod, agnMessBody, agnFileColl, instrFileColl) + " из каталога <" + this.txbx_pathForListDOCfiles.Text + "> \n" + rtxbx_info.Text;
                            }

                            break;
                        }
                        



                        case "IWA":
                            {
                                agnFileColl = GetFiles4(AgnCod, this.txbx_pathForListDOCfiles.Text);

                                if (agnFileColl.Length != 0)
                                {
                                    //agnMail = "oleg.lesnitsky@dhl.com";
                                    //
                                    agnMail = new[] { "dhliwa@list.ru", "ruivoagt@dhl.com", "Alexandr.Pavlov@dhl.com" };

                                                                        rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Для " + AgnCod + SendMail4(agnMail, "Агентские документы PDF  [Агент: " + AgnCod + "] период " + strRepPeriod, agnMessBody, agnFileColl, instrFileColl) + " из каталога <" + this.txbx_pathForListDOCfiles.Text + "> \n" + rtxbx_info.Text;
                                }

                                break;
                            }

                        case "CSY":
                            {
                                agnFileColl = GetFiles4(AgnCod, this.txbx_pathForListDOCfiles.Text);

                                if (agnFileColl.Length != 0)
                                {
                                    //agnMail = "oleg.lesnitsky@dhl.com";
                                    //
                                    agnMail = new[] { "dhl@dhl21.ru", "Timur.Bildanov@dhl.com" };

                                                                        rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Для " + AgnCod + SendMail4(agnMail, "Агентские документы PDF  [Агент: " + AgnCod + "] период " + strRepPeriod, agnMessBody, agnFileColl, instrFileColl) + " из каталога <" + this.txbx_pathForListDOCfiles.Text + "> \n" + rtxbx_info.Text;
                                }

                                break;
                            }

                        case "JOK":
                            {
                                agnFileColl = GetFiles4(AgnCod, this.txbx_pathForListDOCfiles.Text);

                                if (agnFileColl.Length != 0)
                                {
                                    //agnMail = "oleg.lesnitsky@dhl.com";
                                    //
                                    agnMail = new[] { "dhl@dhl21.ru", "Timur.Bildanov@dhl.com" };

                                                                        rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Для " + AgnCod + SendMail4(agnMail, "Агентские документы PDF  [Агент: " + AgnCod + "] период " + strRepPeriod, agnMessBody, agnFileColl, instrFileColl) + " из каталога <" + this.txbx_pathForListDOCfiles.Text + "> \n" + rtxbx_info.Text;
                                }

                                break;
                            }


                        case "CRM":
                            {
                                agnFileColl = GetFiles4(AgnCod, this.txbx_pathForListDOCfiles.Text);

                                if (agnFileColl.Length != 0)
                                {
                                    //agnMail = "oleg.lesnitsky@dhl.com";
                                    //
                                    agnMail = new[] { "buhdelo12@mail.ru", "roman.b.zaitsev@gmail.com", "Denis.Piskunov@dhl.com" };

                                                                        rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Для " + AgnCod + SendMail4(agnMail, "Агентские документы PDF  [Агент: " + AgnCod + "] период " + strRepPeriod, agnMessBody, agnFileColl, instrFileColl) + " из каталога <" + this.txbx_pathForListDOCfiles.Text + "> \n" + rtxbx_info.Text;
                                }

                                break;
                            }

                        case "AER":
                            {
                                agnFileColl = GetFiles4(AgnCod, this.txbx_pathForListDOCfiles.Text);

                                if (agnFileColl.Length != 0)
                                {
                                    //agnMail = "oleg.lesnitsky@dhl.com";
                                    //
                                    agnMail = new[] { "director@dhlsochi.ru", "Denis.Piskunov@dhl.com" };

                                    rtxbx_info.Text = "[" + DateTime.Now.ToString("dd-MM-yyyy HH:mm:ss") + "] Для " + AgnCod + SendMail4(agnMail, "Агентские документы PDF  [Агент: " + AgnCod + "] период " + strRepPeriod, agnMessBody, agnFileColl, instrFileColl) + " из каталога <" + this.txbx_pathForListDOCfiles.Text + "> \n" + rtxbx_info.Text;
                            }

                                break;
                            }

                }

                }

        }

        private void txbx_newPeriod_Click(object sender, EventArgs e)
        {

            if (this.lbl_period.Text != "не опред.")
            { 
                int prepPeriod = Int32.Parse(this.lbl_period.Text) - 1;
                this.txbx_newPeriod.Text = prepPeriod.ToString();
            }

        }
    }
}
