using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutlookAddIn2
{
    public partial class UserControl1 : UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
            usernameBox.Text = Properties.Settings.Default.username;
            passwordBox.Text = Properties.Settings.Default.password;
            securityTokenBox.Text = Properties.Settings.Default.secuityToken;
        }

        private void usernameBox_TextChanged(object sender, EventArgs e)
        {

        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.username = usernameBox.Text;
            Properties.Settings.Default.password = passwordBox.Text;
            Properties.Settings.Default.secuityToken = securityTokenBox.Text;
            Properties.Settings.Default.Save();
            MessageBox.Show("Settings saved.");
            Globals.Ribbons.Ribbon1.login();
        }
    }
}
