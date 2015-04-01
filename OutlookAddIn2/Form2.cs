using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutlookAddIn2
{
    public partial class Form2 : Form
    {
        public String caseNumber = "";
        private String caseID = "";


        public Form2()
        {
            InitializeComponent();
            if (Globals.Ribbons.Ribbon1.login())
            {

            }
            else
            {
                Close();
            }


        }

        private void button1_Click(object sender, EventArgs e)
        {
            caseNumber = textBox1.Text;
            //MessageBox.Show("Case number = " + caseNumber);
            caseID = Globals.Ribbons.Ribbon1.getCaseIDfromCaseNumber(caseNumber);

            if (caseID != "")
            {
                if (Globals.Ribbons.Ribbon1.addEmailToCase(caseID))
                {
                    MessageBox.Show("Email added to case " + caseNumber);
                }
                /*
                if (Globals.Ribbons.Ribbon1.login() && Globals.Ribbons.Ribbon1.addAttachmentToCase(caseID))
                {
                    MessageBox.Show("Attachments added to case " + caseNumber);
                }
                */
            }


            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            caseNumber = "";
            Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
