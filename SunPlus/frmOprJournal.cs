using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SunPlus
{

    public partial class frm_oprJournal : Form
    {
        DAL dal = new DAL();

        DataSet DSJrnal2;
        DataView DVldg, DVldgLad;
        BindingSource bindSource;


        DataSet DSJrnal;
        DataView DVJrnal;


        int TempNum;

        public frm_oprJournal(string InitJrnalNum, string InitJrnalRefer, int InitTempNum, string InitTempText, string lBusUnit)
        {

            InitializeComponent();
            this.Text = this.Text + " Шаблон: <" + InitTempText + ">";

            this.cmbx_journTempl2.SelectedIndex = InitTempNum - 1;
            TempNum = InitTempNum;
            mtxbx_journalNum.Text = InitJrnalNum;
            txbx_referNo.Text = InitJrnalRefer;
            txbx_bunit.Text = lBusUnit;
            txbx_bunit.Enabled = false;

        }

        private void btn_closeFrm_Click(object sender, EventArgs e)
        {
            this.Close();
            //Application.Exit();
        }

        private void btn_getRecords_Click(object sender, EventArgs e)
        {
            string trefer = "";
            string jnumber = "";

            if ((mtxbx_journalNum.Text.ToString() == string.Empty && txbx_referNo.Text.ToString() == string.Empty) ||
                (mtxbx_journalNum.Text.ToString() == "" && txbx_referNo.Text.ToString() == ""))
            {
                MessageBox.Show(" Введите номер журнала и/или номер референса! ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (txbx_bunit.Text.ToString() == string.Empty && txbx_bunit.Text.ToString().Length != 3)
            {
                MessageBox.Show(" Укажите Бизнес-Юнит! ", "  SUN'PLUS message", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
           }

            if (mtxbx_journalNum.Text.ToString() != "")
                jnumber = this.mtxbx_journalNum.Text.Trim();
            else
                jnumber = "%";

            if (txbx_referNo.Text.ToString() != "")
                trefer = this.txbx_referNo.Text.Trim();
            else
                trefer = "%";

            DSJrnal = new DataSet();
            DSJrnal = dal.GetSunJrnal(jnumber, trefer, TempNum, txbx_bunit.Text.ToString());

            DVJrnal = new DataView(DSJrnal.Tables["TabSunJrnalLDG"]);
            if (DVJrnal.Count == 0)
            {
                MessageBox.Show("Записи отсутствуют!");
                return;
            }

            //BindingSource bs = new BindingSource();

            //dgv_jrnlPanel. = DVJrnal;
            //dgv_jrnlPanel.DataSource = DVJrnal;
            dgv_jrnlPanel.DataSource = DSJrnal.Tables[0];
            SettingsDGV_jrnalPanel(TempNum);
        }


        private void SettingsDGV_jrnalPanel(int sTempl)
        {
            try
            {
                
                foreach (DataGridViewColumn dc in dgv_jrnlPanel.Columns)
                {
                    dc.Visible = false;
                    dc.ReadOnly = false;
                }

                dgv_jrnlPanel.Columns["JRNAL_NO"].Visible = true;
                dgv_jrnlPanel.Columns["JRNAL_NO"].DisplayIndex = 1;
                dgv_jrnlPanel.Columns["JRNAL_NO"].HeaderText = "НомерЖурнала";
                
                dgv_jrnlPanel.Columns["JRNAL_LINE"].Visible = true;
                dgv_jrnlPanel.Columns["JRNAL_LINE"].DisplayIndex = 3;
                dgv_jrnlPanel.Columns["JRNAL_LINE"].HeaderText = "Линия";

                switch (sTempl)
                {
                    case 1: 
                        {
                            dgv_jrnlPanel.Columns["TREFERENCE"].Visible = true;
                            dgv_jrnlPanel.Columns["TREFERENCE"].DisplayIndex = 2;
                            dgv_jrnlPanel.Columns["TREFERENCE"].HeaderText = "Референс";

                            dgv_jrnlPanel.Columns["JRNAL_SRCE"].Visible = true;
                            dgv_jrnlPanel.Columns["JRNAL_TYPE"].Visible = true;
                            dgv_jrnlPanel.Columns["TRANS_DATETIME"].Visible = true;
                            dgv_jrnlPanel.Columns["PERIOD"].Visible = true;

                            dgv_jrnlPanel.Columns["ACCNT_CODE"].Visible = true;
                            dgv_jrnlPanel.Columns["ANAL_T6"].Visible = true;
                            dgv_jrnlPanel.Columns["AMOUNT"].Visible = true;
                            dgv_jrnlPanel.Columns["AMOUNT"].DefaultCellStyle.BackColor = Color.LightGreen;
                            dgv_jrnlPanel.Columns["AMOUNT"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                            dgv_jrnlPanel.Columns["OTHER_AMT"].Visible = true;
                            dgv_jrnlPanel.Columns["OTHER_AMT"].DefaultCellStyle.BackColor = Color.Aquamarine;
                            dgv_jrnlPanel.Columns["OTHER_AMT"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                            dgv_jrnlPanel.Columns["D_C"].Visible = true;
                            dgv_jrnlPanel.Columns["DESCRIPTN"].Visible = true;

                            dgv_jrnlPanel.Columns["TRANS_DATETIME"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["TRANS_DATETIME"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;

                            dgv_jrnlPanel.Columns["TREFERENCE"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["TREFERENCE"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;

                            dgv_jrnlPanel.Columns["DESCRIPTN"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["DESCRIPTN"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;

                            //Убрать в последствии
                            dgv_jrnlPanel.Columns["ROUGH_FLAG"].Visible = true;
                            dgv_jrnlPanel.Columns["ALLOCATION"].Visible = true;
                            dgv_jrnlPanel.Columns["ALLOC_IN_PROGRESS"].Visible = true;
                            dgv_jrnlPanel.Columns["IN_USE_FLAG"].Visible = true;

                            break; 
                        }
                    case 2:
                        {
                            dgv_jrnlPanel.Columns["TREFERENCE"].Visible = true;
                            dgv_jrnlPanel.Columns["TREFERENCE"].DisplayIndex = 2;
                            dgv_jrnlPanel.Columns["TREFERENCE"].HeaderText = "Референс";

                            dgv_jrnlPanel.Columns["JRNAL_SRCE"].Visible = true;
                            dgv_jrnlPanel.Columns["JRNAL_TYPE"].Visible = true;
                            dgv_jrnlPanel.Columns["TRANS_DATETIME"].Visible = true;
                            dgv_jrnlPanel.Columns["PERIOD"].Visible = true;

                            dgv_jrnlPanel.Columns["ACCNT_CODE"].Visible = true;
                            dgv_jrnlPanel.Columns["ANAL_T6"].Visible = true;
                            dgv_jrnlPanel.Columns["D_C"].Visible = true;

                            dgv_jrnlPanel.Columns["AMOUNT"].Visible = true;
                            dgv_jrnlPanel.Columns["AMOUNT"].DefaultCellStyle.BackColor = Color.LightGreen;
                            dgv_jrnlPanel.Columns["AMOUNT"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                            dgv_jrnlPanel.Columns["OTHER_AMT"].Visible = true;
                            dgv_jrnlPanel.Columns["OTHER_AMT"].DefaultCellStyle.BackColor = Color.Aquamarine;
                            dgv_jrnlPanel.Columns["OTHER_AMT"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                            dgv_jrnlPanel.Columns["DESCRIPTN"].Visible = true;

                            dgv_jrnlPanel.Columns["TRANS_DATETIME"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["TRANS_DATETIME"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;

                            dgv_jrnlPanel.Columns["TREFERENCE"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["TREFERENCE"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;

                            dgv_jrnlPanel.Columns["DESCRIPTN"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["DESCRIPTN"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;

                            dgv_jrnlPanel.Columns["ROUGH_FLAG"].Visible = true;
                            dgv_jrnlPanel.Columns["ALLOCATION"].Visible = true;
                            dgv_jrnlPanel.Columns["ALLOC_IN_PROGRESS"].Visible = true;
                            dgv_jrnlPanel.Columns["IN_USE_FLAG"].Visible = true;

                            dgv_jrnlPanel.Columns["CONV_CODE"].Visible = true;
                            dgv_jrnlPanel.Columns["CONV_RATE"].Visible = true;

                            dgv_jrnlPanel.Columns["ANAL_T0"].Visible = true;
                            dgv_jrnlPanel.Columns["ANAL_T0"].HeaderText = "T01_STN";
                            dgv_jrnlPanel.Columns["ANAL_T1"].Visible = true;
                            dgv_jrnlPanel.Columns["ANAL_T1"].HeaderText = "T02_PRODT";
                            dgv_jrnlPanel.Columns["ANAL_T2"].Visible = true;
                            dgv_jrnlPanel.Columns["ANAL_T2"].HeaderText = "T03_";
                            dgv_jrnlPanel.Columns["ANAL_T3"].Visible = true;
                            dgv_jrnlPanel.Columns["ANAL_T3"].HeaderText = "T04_PARTN";
                            dgv_jrnlPanel.Columns["ANAL_T4"].Visible = true;
                            dgv_jrnlPanel.Columns["ANAL_T4"].HeaderText = "T05_FUNCT";
                            dgv_jrnlPanel.Columns["ANAL_T5"].Visible = true;
                            dgv_jrnlPanel.Columns["ANAL_T5"].HeaderText = "T06_ASSTYP";
                            dgv_jrnlPanel.Columns["ANAL_T7"].Visible = true;
                            dgv_jrnlPanel.Columns["ANAL_T7"].HeaderText = "T08_ACCCAT";
                            dgv_jrnlPanel.Columns["ANAL_T8"].Visible = true;
                            dgv_jrnlPanel.Columns["ANAL_T8"].HeaderText = "T09_";
                            dgv_jrnlPanel.Columns["ANAL_T9"].Visible = true;
                            dgv_jrnlPanel.Columns["ANAL_T9"].HeaderText = "T10_TAXALL";


                            break; 
                        }
                    case 3:
                        {
                            dgv_jrnlPanel.Columns["TREFERENCE"].Visible = true;
                            dgv_jrnlPanel.Columns["TREFERENCE"].DisplayIndex = 2;
                            dgv_jrnlPanel.Columns["TREFERENCE"].HeaderText = "Референс";

                            dgv_jrnlPanel.Columns["GD02_NOMER_FAKTURA"].Visible = true;
                            dgv_jrnlPanel.Columns["GD02_NOMER_FAKTURA"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["GD02_NOMER_FAKTURA"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                            dgv_jrnlPanel.Columns["GD02_NOMER_FAKTURA"].DisplayIndex = 4;
                            dgv_jrnlPanel.Columns["GD02_NOMER_FAKTURA"].HeaderText = "Номер с/ф";

                            dgv_jrnlPanel.Columns["GD5_OLD_JNL_TYP"].Visible = true;
                            dgv_jrnlPanel.Columns["GD5_OLD_JNL_TYP"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["GD5_OLD_JNL_TYP"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                            dgv_jrnlPanel.Columns["GD5_OLD_JNL_TYP"].DisplayIndex = 5;
                            dgv_jrnlPanel.Columns["GD5_OLD_JNL_TYP"].HeaderText = "Old Journal Type";

                            dgv_jrnlPanel.Columns["GDT3_DATE_FAKTURA"].Visible = true;
                            dgv_jrnlPanel.Columns["GDT3_DATE_FAKTURA"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["GDT3_DATE_FAKTURA"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                            dgv_jrnlPanel.Columns["GDT3_DATE_FAKTURA"].DisplayIndex = 6;
                            dgv_jrnlPanel.Columns["GDT3_DATE_FAKTURA"].HeaderText = "Дата с/ф";

                            dgv_jrnlPanel.Columns["GDT1_DATE_OPLATA"].Visible = true;
                            dgv_jrnlPanel.Columns["GDT1_DATE_OPLATA"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["GDT1_DATE_OPLATA"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                            dgv_jrnlPanel.Columns["GDT1_DATE_OPLATA"].DisplayIndex = 7;
                            dgv_jrnlPanel.Columns["GDT1_DATE_OPLATA"].HeaderText = "ДатаОплаты с/ф";

                            dgv_jrnlPanel.Columns["GDT2_DATE_PRIN_UCHET"].Visible = true;
                            dgv_jrnlPanel.Columns["GDT2_DATE_PRIN_UCHET"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["GDT2_DATE_PRIN_UCHET"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                            dgv_jrnlPanel.Columns["GDT2_DATE_PRIN_UCHET"].DisplayIndex = 8;
                            dgv_jrnlPanel.Columns["GDT2_DATE_PRIN_UCHET"].HeaderText = "ДатаПринятияНаУчет";

                            dgv_jrnlPanel.Columns["GD11_KODOPLATY"].Visible = true;
                            dgv_jrnlPanel.Columns["GD11_KODOPLATY"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["GD11_KODOPLATY"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                            dgv_jrnlPanel.Columns["GD11_KODOPLATY"].DisplayIndex = 9;
                            dgv_jrnlPanel.Columns["GD11_KODOPLATY"].HeaderText = "КодОплаты";

                            dgv_jrnlPanel.Columns["JRNAL_SRCE"].Visible = true;
                            dgv_jrnlPanel.Columns["JRNAL_SRCE"].DisplayIndex = 10;
                            dgv_jrnlPanel.Columns["JRNAL_SRCE"].HeaderText = "Пользователь";

                            dgv_jrnlPanel.Columns["TRANS_DATETIME"].Visible = true;
                            dgv_jrnlPanel.Columns["TRANS_DATETIME"].DisplayIndex = 11;
                            dgv_jrnlPanel.Columns["TRANS_DATETIME"].HeaderText = "ДатаТранзакции";

                            dgv_jrnlPanel.Columns["PERIOD"].Visible = true;
                            dgv_jrnlPanel.Columns["PERIOD"].DisplayIndex = 12;
                            dgv_jrnlPanel.Columns["PERIOD"].HeaderText = "Период";

                            dgv_jrnlPanel.Columns["ACCNT_CODE"].Visible = true;
                            dgv_jrnlPanel.Columns["ACCNT_CODE"].DisplayIndex = 13;
                            dgv_jrnlPanel.Columns["ACCNT_CODE"].HeaderText = "Счет";

                            dgv_jrnlPanel.Columns["ANAL_T6"].Visible = true;
                            dgv_jrnlPanel.Columns["ANAL_T6"].DisplayIndex = 14;
                            dgv_jrnlPanel.Columns["ANAL_T6"].HeaderText = "КорСчет";

                            dgv_jrnlPanel.Columns["AMOUNT"].Visible = true;
                            dgv_jrnlPanel.Columns["AMOUNT"].DisplayIndex = 15;
                            dgv_jrnlPanel.Columns["AMOUNT"].HeaderText = "Сумма";

                            dgv_jrnlPanel.Columns["D_C"].Visible = true;
                            dgv_jrnlPanel.Columns["D_C"].DisplayIndex = 16;
                            dgv_jrnlPanel.Columns["D_C"].HeaderText = "Д/К";

                            dgv_jrnlPanel.Columns["DESCRIPTN"].Visible = true;
                            dgv_jrnlPanel.Columns["DESCRIPTN"].DisplayIndex = 17;
                            dgv_jrnlPanel.Columns["DESCRIPTN"].HeaderText = "Description";

                            break;
                        }
                    case 4:
                        {

                            dgv_jrnlPanel.Columns["TREFERENCE"].Visible = true;
                            dgv_jrnlPanel.Columns["TREFERENCE"].DisplayIndex = 2;
                            dgv_jrnlPanel.Columns["TREFERENCE"].HeaderText = "Референс";

                            dgv_jrnlPanel.Columns["GD02_NOMER_FAKTURA"].Visible = true;
                            dgv_jrnlPanel.Columns["GD02_NOMER_FAKTURA"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["GD02_NOMER_FAKTURA"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                            dgv_jrnlPanel.Columns["GD02_NOMER_FAKTURA"].DisplayIndex = 4;
                            dgv_jrnlPanel.Columns["GD02_NOMER_FAKTURA"].HeaderText = "Номер с/ф";

                            dgv_jrnlPanel.Columns["GD5_OLD_JNL_TYP"].Visible = true;
                            dgv_jrnlPanel.Columns["GD5_OLD_JNL_TYP"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["GD5_OLD_JNL_TYP"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                            dgv_jrnlPanel.Columns["GD5_OLD_JNL_TYP"].DisplayIndex = 5;
                            dgv_jrnlPanel.Columns["GD5_OLD_JNL_TYP"].HeaderText = "Old Journal Type";

                            dgv_jrnlPanel.Columns["GDT3_DATE_FAKTURA"].Visible = true;
                            dgv_jrnlPanel.Columns["GDT3_DATE_FAKTURA"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["GDT3_DATE_FAKTURA"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                            dgv_jrnlPanel.Columns["GDT3_DATE_FAKTURA"].DisplayIndex = 6;
                            dgv_jrnlPanel.Columns["GDT3_DATE_FAKTURA"].HeaderText = "Дата с/ф";

                            dgv_jrnlPanel.Columns["GDT1_DATE_OPLATA"].Visible = true;
                            dgv_jrnlPanel.Columns["GDT1_DATE_OPLATA"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["GDT1_DATE_OPLATA"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                            dgv_jrnlPanel.Columns["GDT1_DATE_OPLATA"].DisplayIndex = 7;
                            dgv_jrnlPanel.Columns["GDT1_DATE_OPLATA"].HeaderText = "ДатаОплаты с/ф";

                            dgv_jrnlPanel.Columns["GDT2_DATE_PRIN_UCHET"].Visible = true;
                            dgv_jrnlPanel.Columns["GDT2_DATE_PRIN_UCHET"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["GDT2_DATE_PRIN_UCHET"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                            dgv_jrnlPanel.Columns["GDT2_DATE_PRIN_UCHET"].DisplayIndex = 8;
                            dgv_jrnlPanel.Columns["GDT2_DATE_PRIN_UCHET"].HeaderText = "ДатаПринятияНаУчет";

                            dgv_jrnlPanel.Columns["GD11_KODOPLATY"].Visible = true;
                            dgv_jrnlPanel.Columns["GD11_KODOPLATY"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["GD11_KODOPLATY"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                            dgv_jrnlPanel.Columns["GD11_KODOPLATY"].DisplayIndex = 9;
                            dgv_jrnlPanel.Columns["GD11_KODOPLATY"].HeaderText = "КодОплаты";

                            dgv_jrnlPanel.Columns["JRNAL_SRCE"].Visible = true;
                            dgv_jrnlPanel.Columns["JRNAL_SRCE"].DisplayIndex = 10;
                            dgv_jrnlPanel.Columns["JRNAL_SRCE"].HeaderText = "Пользователь";

                            dgv_jrnlPanel.Columns["TRANS_DATETIME"].Visible = true;
                            dgv_jrnlPanel.Columns["TRANS_DATETIME"].DisplayIndex = 11;
                            dgv_jrnlPanel.Columns["TRANS_DATETIME"].HeaderText = "ДатаТранзакции";

                            dgv_jrnlPanel.Columns["PERIOD"].Visible = true;
                            dgv_jrnlPanel.Columns["PERIOD"].DisplayIndex = 12;
                            dgv_jrnlPanel.Columns["PERIOD"].HeaderText = "Период";

                            dgv_jrnlPanel.Columns["ACCNT_CODE"].Visible = true;
                            dgv_jrnlPanel.Columns["ACCNT_CODE"].DisplayIndex = 13;
                            dgv_jrnlPanel.Columns["ACCNT_CODE"].HeaderText = "Счет";

                            dgv_jrnlPanel.Columns["ANAL_T6"].Visible = true;
                            dgv_jrnlPanel.Columns["ANAL_T6"].DisplayIndex = 14;
                            dgv_jrnlPanel.Columns["ANAL_T6"].HeaderText = "КорСчет";

                            dgv_jrnlPanel.Columns["AMOUNT"].Visible = true;
                            dgv_jrnlPanel.Columns["AMOUNT"].DisplayIndex = 15;
                            dgv_jrnlPanel.Columns["AMOUNT"].HeaderText = "Сумма";

                            dgv_jrnlPanel.Columns["D_C"].Visible = true;
                            dgv_jrnlPanel.Columns["D_C"].DisplayIndex = 16;
                            dgv_jrnlPanel.Columns["D_C"].HeaderText = "Д/К";

                            dgv_jrnlPanel.Columns["DESCRIPTN"].Visible = true;
                            dgv_jrnlPanel.Columns["DESCRIPTN"].DisplayIndex = 17;
                            dgv_jrnlPanel.Columns["DESCRIPTN"].HeaderText = "Description";

                            break;
                        }
                    case 5:
                        {
                            dgv_jrnlPanel.Columns["GD02_NOMER_FAKTURA"].Visible = true;
                            dgv_jrnlPanel.Columns["GD02_NOMER_FAKTURA"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["GD02_NOMER_FAKTURA"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                            dgv_jrnlPanel.Columns["GD02_NOMER_FAKTURA"].DisplayIndex = 4;
                            dgv_jrnlPanel.Columns["GD02_NOMER_FAKTURA"].HeaderText = "Номер с/ф";

                            dgv_jrnlPanel.Columns["GD5_OLD_JNL_TYP"].Visible = true;
                            dgv_jrnlPanel.Columns["GD5_OLD_JNL_TYP"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["GD5_OLD_JNL_TYP"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                            dgv_jrnlPanel.Columns["GD5_OLD_JNL_TYP"].DisplayIndex = 5;
                            dgv_jrnlPanel.Columns["GD5_OLD_JNL_TYP"].HeaderText = "Old Journal Type";

                            dgv_jrnlPanel.Columns["GDT3_DATE_FAKTURA"].Visible = true;
                            dgv_jrnlPanel.Columns["GDT3_DATE_FAKTURA"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["GDT3_DATE_FAKTURA"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                            dgv_jrnlPanel.Columns["GDT3_DATE_FAKTURA"].DisplayIndex = 6;
                            dgv_jrnlPanel.Columns["GDT3_DATE_FAKTURA"].HeaderText = "Дата с/ф";

                            dgv_jrnlPanel.Columns["GDT1_DATE_OPLATA"].Visible = true;
                            dgv_jrnlPanel.Columns["GDT1_DATE_OPLATA"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["GDT1_DATE_OPLATA"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                            dgv_jrnlPanel.Columns["GDT1_DATE_OPLATA"].DisplayIndex = 7;
                            dgv_jrnlPanel.Columns["GDT1_DATE_OPLATA"].HeaderText = "ДатаОплаты с/ф";

                            dgv_jrnlPanel.Columns["GDT2_DATE_PRIN_UCHET"].Visible = true;
                            dgv_jrnlPanel.Columns["GDT2_DATE_PRIN_UCHET"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["GDT2_DATE_PRIN_UCHET"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                            dgv_jrnlPanel.Columns["GDT2_DATE_PRIN_UCHET"].DisplayIndex = 8;
                            dgv_jrnlPanel.Columns["GDT2_DATE_PRIN_UCHET"].HeaderText = "ДатаПринятияНаУчет";

                            dgv_jrnlPanel.Columns["GD11_KODOPLATY"].Visible = true;
                            dgv_jrnlPanel.Columns["GD11_KODOPLATY"].ReadOnly = false;
                            dgv_jrnlPanel.Columns["GD11_KODOPLATY"].DefaultCellStyle.BackColor = Color.PaleGoldenrod;
                            dgv_jrnlPanel.Columns["GD11_KODOPLATY"].DisplayIndex = 9;
                            dgv_jrnlPanel.Columns["GD11_KODOPLATY"].HeaderText = "КодОплаты";


                            dgv_jrnlPanel.Columns["PERIOD"].Visible = true;
                            dgv_jrnlPanel.Columns["PERIOD"].DisplayIndex = 12;
                            dgv_jrnlPanel.Columns["PERIOD"].HeaderText = "Период";

                            break;
                        }
                    case 6:
                            {
                            foreach (DataGridViewColumn dc in dgv_jrnlPanel.Columns)
                                    {
                                        dc.Visible = true;
                                        dc.ReadOnly = true;
                                    }
                            break;
                            }

                }

            }
            catch
            { }



        }

        private void cbtn_changeRecords_Click(object sender, EventArgs e)
        {

            //DAL.UpdateJrnalLDG(DSJrnal.Tables["TabSunJrnalLDG"]);

            dgv_jrnlPanel.DataSource = dal.UpdateJrnalLDG().Tables[0];
            this.Text = this.Text.Replace("*","!");
        }

        private void dgv_jrnlPanel_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            this.Text = this.Text + "***";
        }

        private void cbtn_clear1_Click(object sender, EventArgs e)
        {
            mtxbx_journalNum.Text = "";
            txbx_referNo.Text = "";
            mtxbx_journalNum.Focus();
        }

    }

    public class cmbx_ColumnSet : ComboBox
    {

        public cmbx_ColumnSet()
        {
            this.FormattingEnabled = true;
            Items.AddRange(new object[] {
            "1. Журнал [LDG] v.1",
            "11. Журнал простой [trans_datetime] v.1",
            "12. Журнал простой [descriptn] v.1",
            "13. Журнал с просмотром С/Ф v.1",
            "14. Журнал [просмотр] полный",
            "15. Другой журнал2",
            "16. Системный журнал [Все ячейки]"});
        }

    }

}
