using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Xml;
using System.Xml.XPath;


namespace SunPlus
{
    public partial class frm_SunXMLprotocol : Form
    {

        public frm_SunXMLprotocol(ref DataTable XMLrows)
        {
            InitializeComponent();
            
            dgv_xmlData.DataSource = XMLrows;

            if (XMLrows.Rows.Count == 0)
                lbl_XMLProtocolResult.Text = "Строки не найдены.";
                else
                lbl_XMLProtocolResult.Text = "Сформировано " + XMLrows.Rows.Count.ToString() + " строк.";

        }

        private void btn_ReadXML_Click(object sender, EventArgs e)
        {
            string myXMLfile = @"C:\temp\SSCLog3228428084274704391.xml";
            DataSet ds = new DataSet();
            // Create new FileStream with which to read the schema.
            System.IO.FileStream fsReadXml = new System.IO.FileStream
                (myXMLfile, System.IO.FileMode.Open);
            try
            {
                ds.ReadXml(fsReadXml);
                dgv_xmlData.DataSource = ds;
                //dgv_xmlData.DataMember = "TransferDescSunProtocol";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                fsReadXml.Close();
            }
        }

        private void btn_File_Click(object sender, EventArgs e)
        {
            string myXMLfile = "C:\\temp\\SSCLog3228428084274704391.xml";
            DataSet ds = new DataSet();
            try
            {
                ds.ReadXml(myXMLfile);
                dgv_xmlData.DataSource = ds;
                //dgv_xmlData.DataMember = "TransferDescSunProtocol";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btn_closeFrm2_Click_Click(object sender, EventArgs e)
        {

            this.Close();

        }

        private void frm_SunXMLprotocol_Load(object sender, EventArgs e)
        {

        }
    }
}
