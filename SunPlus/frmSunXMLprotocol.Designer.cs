namespace SunPlus
{
    partial class frm_SunXMLprotocol
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frm_SunXMLprotocol));
            this.dgv_xmlData = new System.Windows.Forms.DataGridView();
            this.btn_closeFrm2_Click = new System.Windows.Forms.Button();
            this.lbl_XMLProtocolResult = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_xmlData)).BeginInit();
            this.SuspendLayout();
            // 
            // dgv_xmlData
            // 
            this.dgv_xmlData.AllowUserToAddRows = false;
            this.dgv_xmlData.AllowUserToDeleteRows = false;
            this.dgv_xmlData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv_xmlData.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells;
            this.dgv_xmlData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_xmlData.Location = new System.Drawing.Point(25, 35);
            this.dgv_xmlData.Name = "dgv_xmlData";
            this.dgv_xmlData.ReadOnly = true;
            this.dgv_xmlData.Size = new System.Drawing.Size(1050, 171);
            this.dgv_xmlData.TabIndex = 0;
            // 
            // btn_closeFrm2_Click
            // 
            this.btn_closeFrm2_Click.Location = new System.Drawing.Point(957, 240);
            this.btn_closeFrm2_Click.Name = "btn_closeFrm2_Click";
            this.btn_closeFrm2_Click.Size = new System.Drawing.Size(105, 31);
            this.btn_closeFrm2_Click.TabIndex = 1;
            this.btn_closeFrm2_Click.Text = "Close";
            this.btn_closeFrm2_Click.UseVisualStyleBackColor = true;
            this.btn_closeFrm2_Click.Click += new System.EventHandler(this.btn_closeFrm2_Click_Click);
            // 
            // lbl_XMLProtocolResult
            // 
            this.lbl_XMLProtocolResult.AutoSize = true;
            this.lbl_XMLProtocolResult.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.lbl_XMLProtocolResult.ForeColor = System.Drawing.SystemColors.Desktop;
            this.lbl_XMLProtocolResult.Location = new System.Drawing.Point(164, 228);
            this.lbl_XMLProtocolResult.Name = "lbl_XMLProtocolResult";
            this.lbl_XMLProtocolResult.Size = new System.Drawing.Size(41, 13);
            this.lbl_XMLProtocolResult.TabIndex = 2;
            this.lbl_XMLProtocolResult.Text = "label1";
            // 
            // frm_SunXMLprotocol
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1098, 294);
            this.Controls.Add(this.lbl_XMLProtocolResult);
            this.Controls.Add(this.btn_closeFrm2_Click);
            this.Controls.Add(this.dgv_xmlData);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximumSize = new System.Drawing.Size(1114, 333);
            this.MinimumSize = new System.Drawing.Size(1114, 333);
            this.Name = "frm_SunXMLprotocol";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "  Протокол ошибок  SunSystems [TransferDesk]";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.frm_SunXMLprotocol_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv_xmlData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgv_xmlData;
        private System.Windows.Forms.Button btn_closeFrm2_Click;
        private System.Windows.Forms.Label lbl_XMLProtocolResult;
    }
}