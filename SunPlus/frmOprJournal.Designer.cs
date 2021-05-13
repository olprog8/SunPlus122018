namespace SunPlus
{
    partial class frm_oprJournal
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frm_oprJournal));
            this.txbx_referNo = new System.Windows.Forms.TextBox();
            this.btn_getRecords = new System.Windows.Forms.Button();
            this.mtxbx_journalNum = new System.Windows.Forms.MaskedTextBox();
            this.btn_closeFrm = new System.Windows.Forms.Button();
            this.dgv_jrnlPanel = new System.Windows.Forms.DataGridView();
            this.label3 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.cbtn_changeRecords = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.cmbx_journTempl2 = new SunPlus.cmbx_ColumnSet();
            this.label2 = new System.Windows.Forms.Label();
            this.txbx_bunit = new System.Windows.Forms.TextBox();
            this.cbtn_clear1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_jrnlPanel)).BeginInit();
            this.SuspendLayout();
            // 
            // txbx_referNo
            // 
            this.txbx_referNo.Font = new System.Drawing.Font("Arial Rounded MT Bold", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txbx_referNo.Location = new System.Drawing.Point(668, 39);
            this.txbx_referNo.MaxLength = 50;
            this.txbx_referNo.Name = "txbx_referNo";
            this.txbx_referNo.Size = new System.Drawing.Size(267, 23);
            this.txbx_referNo.TabIndex = 2;
            this.txbx_referNo.Text = "140699";
            this.txbx_referNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // btn_getRecords
            // 
            this.btn_getRecords.Location = new System.Drawing.Point(953, 35);
            this.btn_getRecords.Name = "btn_getRecords";
            this.btn_getRecords.Size = new System.Drawing.Size(105, 29);
            this.btn_getRecords.TabIndex = 3;
            this.btn_getRecords.Text = "Load";
            this.btn_getRecords.UseVisualStyleBackColor = true;
            this.btn_getRecords.Click += new System.EventHandler(this.btn_getRecords_Click);
            // 
            // mtxbx_journalNum
            // 
            this.mtxbx_journalNum.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.mtxbx_journalNum.Location = new System.Drawing.Point(564, 38);
            this.mtxbx_journalNum.Mask = "000000";
            this.mtxbx_journalNum.Name = "mtxbx_journalNum";
            this.mtxbx_journalNum.Size = new System.Drawing.Size(84, 25);
            this.mtxbx_journalNum.TabIndex = 4;
            this.mtxbx_journalNum.Text = "201348";
            this.mtxbx_journalNum.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // btn_closeFrm
            // 
            this.btn_closeFrm.Location = new System.Drawing.Point(953, 342);
            this.btn_closeFrm.Name = "btn_closeFrm";
            this.btn_closeFrm.Size = new System.Drawing.Size(105, 31);
            this.btn_closeFrm.TabIndex = 5;
            this.btn_closeFrm.Text = "Close";
            this.btn_closeFrm.UseVisualStyleBackColor = true;
            this.btn_closeFrm.Click += new System.EventHandler(this.btn_closeFrm_Click);
            // 
            // dgv_jrnlPanel
            // 
            this.dgv_jrnlPanel.AllowUserToAddRows = false;
            this.dgv_jrnlPanel.AllowUserToDeleteRows = false;
            this.dgv_jrnlPanel.AllowUserToOrderColumns = true;
            this.dgv_jrnlPanel.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Tahoma", 6.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgv_jrnlPanel.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv_jrnlPanel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Tahoma", 8F, System.Drawing.FontStyle.Bold);
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv_jrnlPanel.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgv_jrnlPanel.Location = new System.Drawing.Point(12, 85);
            this.dgv_jrnlPanel.Name = "dgv_jrnlPanel";
            this.dgv_jrnlPanel.Size = new System.Drawing.Size(1066, 174);
            this.dgv_jrnlPanel.TabIndex = 6;
            this.dgv_jrnlPanel.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv_jrnlPanel_CellEndEdit);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(561, 21);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(87, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Номер журнала";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(665, 23);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(58, 13);
            this.label9.TabIndex = 8;
            this.label9.Text = "Референс";
            // 
            // cbtn_changeRecords
            // 
            this.cbtn_changeRecords.Location = new System.Drawing.Point(964, 293);
            this.cbtn_changeRecords.Name = "cbtn_changeRecords";
            this.cbtn_changeRecords.Size = new System.Drawing.Size(94, 33);
            this.cbtn_changeRecords.TabIndex = 9;
            this.cbtn_changeRecords.Text = "Amend";
            this.cbtn_changeRecords.UseVisualStyleBackColor = true;
            this.cbtn_changeRecords.Click += new System.EventHandler(this.cbtn_changeRecords_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(24, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(90, 13);
            this.label1.TabIndex = 10;
            this.label1.Text = "Шаблон вывода:";
            // 
            // cmbx_journTempl2
            // 
            this.cmbx_journTempl2.FormattingEnabled = true;
            this.cmbx_journTempl2.Location = new System.Drawing.Point(27, 40);
            this.cmbx_journTempl2.Name = "cmbx_journTempl2";
            this.cmbx_journTempl2.Size = new System.Drawing.Size(373, 21);
            this.cmbx_journTempl2.TabIndex = 11;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(424, 21);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(73, 13);
            this.label2.TabIndex = 12;
            this.label2.Text = "Бизнес-Юнит";
            // 
            // txbx_bunit
            // 
            this.txbx_bunit.Location = new System.Drawing.Point(427, 38);
            this.txbx_bunit.Name = "txbx_bunit";
            this.txbx_bunit.Size = new System.Drawing.Size(60, 20);
            this.txbx_bunit.TabIndex = 13;
            // 
            // cbtn_clear1
            // 
            this.cbtn_clear1.Location = new System.Drawing.Point(909, 12);
            this.cbtn_clear1.Name = "cbtn_clear1";
            this.cbtn_clear1.Size = new System.Drawing.Size(25, 22);
            this.cbtn_clear1.TabIndex = 14;
            this.cbtn_clear1.Text = "C";
            this.cbtn_clear1.UseVisualStyleBackColor = true;
            this.cbtn_clear1.Click += new System.EventHandler(this.cbtn_clear1_Click);
            // 
            // frm_oprJournal
            // 
            this.ClientSize = new System.Drawing.Size(1086, 388);
            this.Controls.Add(this.cbtn_clear1);
            this.Controls.Add(this.txbx_bunit);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cmbx_journTempl2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbtn_changeRecords);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.dgv_jrnlPanel);
            this.Controls.Add(this.btn_closeFrm);
            this.Controls.Add(this.mtxbx_journalNum);
            this.Controls.Add(this.btn_getRecords);
            this.Controls.Add(this.txbx_referNo);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frm_oprJournal";
            this.Text = "  SUN\'PLUS Просмотр журнала.";
            this.Load += new System.EventHandler(this.btn_getRecords_Click);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_jrnlPanel)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txbx_referNo;
        private System.Windows.Forms.Button btn_getRecords;
        private System.Windows.Forms.MaskedTextBox mtxbx_journalNum;
        private System.Windows.Forms.Button btn_closeFrm;
        private System.Windows.Forms.DataGridView dgv_jrnlPanel;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Button cbtn_changeRecords;
        private System.Windows.Forms.Label label1;
        private cmbx_ColumnSet cmbx_journTempl2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txbx_bunit;
        private System.Windows.Forms.Button cbtn_clear1;
//       private cmbx_ColumnSet cmbx_journTempl2;
    }
}