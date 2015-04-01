using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using OutlookAddIn2.WebReference;
using System.Windows.Forms;

namespace OutlookAddIn2
{
    public partial class Ribbon1
    {
        private SforceService sfs = null;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            if (Properties.Settings.Default.username != "")
            {
                login();
            }
        }

        private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        {
            //sets pane on/off function to toggle button
            Globals.ThisAddIn.customPane.Visible = !Globals.ThisAddIn.customPane.Visible;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 f = new Form1();
            f.Show();
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Form2 f = new Form2();
            f.Show();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            if (login())
            {
                MessageBox.Show("Login Successful");
            }
        }



        public bool login()
        {
            LoginResult lr;
            try
            {
                sfs = new SforceService();
                sfs.Timeout = 5000;
                lr = sfs.login(Properties.Settings.Default.username, Properties.Settings.Default.password + Properties.Settings.Default.secuityToken);
                sfs.Url = lr.serverUrl;
                sfs.SessionHeaderValue = new SessionHeader();
                sfs.SessionHeaderValue.sessionId = lr.sessionId;
            }
            catch (System.Web.Services.Protocols.SoapException e)
            {
                // This is likley to be caused by bad username or password
                MessageBox.Show(e.Message + ", please try again.\n\nHit return to continue...");
                return false;
            }
            catch (Exception e)
            {
                // This is something else, probably comminication
                MessageBox.Show(e.Message + ", please try again.\n\nHit return to continue...");
                return false;
            }
            //MessageBox.Show("Login Successful");
            return true;
        }

        public string getCaseIDfromCaseNumber(string caseNumber)
        {
            SearchResult sr = null;

            try
            {
                sr = sfs.search("FIND {" + caseNumber + "} IN ALL FIELDS RETURNING CASE(Id, ContactId)");
                //MessageBox.Show(sr.searchRecords.Length.ToString());
                if (sr != null)
                {
                    if (sr.searchRecords != null)
                    {
                        SearchRecord[] records = sr.searchRecords;
                        //MessageBox.Show("There are " + records.Length.ToString() + " matches");
                        if (records.Length == 1)
                        {
                            return records[0].record.Id;
                        }
                        else if (records.Length > 1)
                        {
                            MessageBox.Show("Error: Too many CaseIDs found");
                        }
                        else
                        {
                            MessageBox.Show("Error: No records found");
                        }
                    }
                    else
                    {
                        MessageBox.Show("Error: No records found");
                    }
                }
                else
                {
                    MessageBox.Show("Error: sr is null");
                }

            }

            catch (Exception e)
            {
                MessageBox.Show("Error: " + e.Message);
            }
            return "";
        }

        public bool addAttachmentToCase(string caseID)
        {
            Attachment att = new Attachment();

            System.IO.FileInfo attachment;
            System.IO.FileStream fs;
            byte[] bytes;

            SaveResult[] saveResults;

            if (Globals.ThisAddIn.currentMailItem.Attachments != null && Globals.ThisAddIn.currentMailItem.Attachments.Count > 0)
            {
                //MessageBox.Show("Number of attachments = " + Globals.ThisAddIn.currentMailItem.Attachments.Count);

                string tempFilePath;

                for (int i = 0; i < Globals.ThisAddIn.currentMailItem.Attachments.Count; i++)
                {
                    
                    try
                    {
                        //create a temp file and save the outlook attachment there
                        tempFilePath = System.IO.Path.GetTempFileName();
                        //attachments array starts at 1 instead of 0
                        Globals.ThisAddIn.currentMailItem.Attachments[i + 1].SaveAsFile(tempFilePath);
                    }
                    catch (System.IO.IOException e)
                    {
                        MessageBox.Show(e.Message);
                        return false;
                    }

                    //First get a reference to the file you want to send
                    //attachment = new System.IO.FileInfo(Globals.ThisAddIn.currentMailItem.Attachments[i + 1].GetTemporaryFilePath());
                    attachment = new System.IO.FileInfo(tempFilePath);

                    if (attachment.Length >= 5242880)
                    {
                        MessageBox.Show("Error: Attachment over 5 MB");
                        return false;
                    }

                    //Open the file for binary read operation
                    fs = attachment.OpenRead();

                    //create a variable to hold the binary (byte) data
                    bytes = new byte[fs.Length];

                    //read the file into the byte array in it's entirety
                    fs.Read(bytes, 0, (int)fs.Length);

                    //close the file
                    fs.Close();

                    att.Body = bytes;
                    att.Name = Globals.ThisAddIn.currentMailItem.Attachments[i + 1].FileName;
                    att.IsPrivate = false;
                    att.ParentId = caseID;

                    saveResults = sfs.create(new sObject[] { att });

                    if (saveResults[0].success)
                    {
                        //MessageBox.Show("An account with an id of: " + saveResults[i].id + " was updated.\n");
                    }
                    else
                    {
                        MessageBox.Show("Item " + i + " had an error updating.");
                        MessageBox.Show("  The error reported was: " + saveResults[0].errors[0].message + "\n");
                    }
                }

            }
            else
            {
                //no attachments
                MessageBox.Show("Error: No attachments in email");
                return false;
            }

            return true;
        }

        public bool addEmailToCase(string caseID)
        {
            Attachment att = new Attachment();

            System.IO.FileInfo attachment;
            System.IO.FileStream fs;
            byte[] bytes;

            SaveResult[] saveResults;

            string tempFilePath;

            try
            {
                //create a temp file and save the outlook attachment there
                tempFilePath = System.IO.Path.GetTempFileName();
                //attachments array starts at 1 instead of 0
                //Globals.ThisAddIn.currentMailItem.Attachments[i + 1].SaveAsFile(tempFilePath);
                Globals.ThisAddIn.currentMailItem.SaveAs(tempFilePath);
            }
            catch (System.IO.IOException e)
            {
                MessageBox.Show(e.Message);
                return false;
            }
            //First get a reference to the file you want to send
            //attachment = new System.IO.FileInfo(Globals.ThisAddIn.currentMailItem.Attachments[i + 1].GetTemporaryFilePath());
            attachment = new System.IO.FileInfo(tempFilePath);

            if (attachment.Length >= 5242880)
            {
                MessageBox.Show("Error: Email over 5 MB");
                return false;
            }

            //Open the file for binary read operation
            fs = attachment.OpenRead();

            //create a variable to hold the binary (byte) data
            bytes = new byte[fs.Length];

            //read the file into the byte array in it's entirety
            fs.Read(bytes, 0, (int)fs.Length);

            //close the file
            fs.Close();

            att.Body = bytes;
            att.Name = Globals.ThisAddIn.currentMailItem.Subject + ".msg";
            att.IsPrivate = false;
            att.ParentId = caseID;

            saveResults = sfs.create(new sObject[] { att });

            if (saveResults[0].success)
            {
                //MessageBox.Show("An account with an id of: " + saveResults[i].id + " was updated.\n");
            }
            else
            {
                //MessageBox.Show("Item " + i + " had an error updating.");
                MessageBox.Show("  The error reported was: " + saveResults[0].errors[0].message + "\n");
            }

            return true;
        }





    }
}
