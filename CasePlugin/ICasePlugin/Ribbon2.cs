using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.IO;

namespace ICasePlugin
{
    public partial class Ribbon2
    {

        private void Ribbon2_Load(object sender, RibbonUIEventArgs e)
        {
            this.GetCase();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            using (FormSettings formSettings = new FormSettings())
            {
                formSettings.ShowDialog();
            }

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (comboBox1.Text == "")
            {
                MessageBox.Show("Please, select the case number");
                return;
            }

            Outlook.MailItem mailItem = null;

            try
            {
                mailItem = (Outlook.MailItem)Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
            }
            catch
            {
                return;
            }

            XCaseService.MessageClass messageClass = new XCaseService.MessageClass();
            messageClass.Body = mailItem.Body;
            messageClass.DateReceived = mailItem.ReceivedTime;
            messageClass.DateSent = mailItem.SentOn;
            messageClass.EntryID = mailItem.EntryID;

            string intCaseNumber = "";
            if (comboBox1.Text.Contains(":"))
            {
                intCaseNumber = comboBox1.Text.Substring(comboBox1.Text.IndexOf(":") + 2);
            }
            else
            {
                intCaseNumber = comboBox1.Text;
            }

            messageClass.CaseID = intCaseNumber.Trim(); 
            
            foreach (Outlook.Recipient recipient in mailItem.Recipients)
            {
                messageClass.Recipient += recipient.Address + ";";
            }

            messageClass.Sender = mailItem.SenderEmailAddress;
            messageClass.Subject = mailItem.Subject;

            if (mailItem.Attachments.Count > 0)
            {
                messageClass.Attachment = true;
            }
            else
            {
                messageClass.Attachment = false;
            }

            // Save to folder
            string fileName = DateTime.Now.ToString("MMddyyyy-hhmmss") + ".msg";
            mailItem.SaveAs(classService.MSGPath + fileName, Outlook.OlSaveAsType.olMSG);

            // Load as byte array 
            System.IO.FileInfo fi = new System.IO.FileInfo(classService.MSGPath + fileName);
            using (System.IO.FileStream fs = fi.Open(System.IO.FileMode.Open,
                                                     System.IO.FileAccess.Read))
            {
                messageClass.Message = new byte[fs.Length];
                int readBytes = fs.Read(messageClass.Message, 0, (int)fs.Length);
                fs.Close();
            }
            fi = null;

            XCaseService.CaseService caseService = new XCaseService.CaseService();
            caseService.Url = classService.URL;
            //if (caseService.Login(classService.Username, classService.Password))
            //{
            //    MessageBox.Show("Invalid credential. Please check on configuration");
            //    return;
            //}

            //MessageBox.Show(caseService.Uploads(messageClass, classService.Username));
            caseService.Uploads(messageClass, classService.Username);
            caseService.Dispose();
            caseService = null;

            try
            {
                File.Delete(fileName);
            }
            catch
            { }

            this.GetCase();

        }

        #region "Methods, Variables"
        ClassService classService = new ClassService();

        private void GetCase()
        {
            try
            {
                XCaseService.CaseService service = new XCaseService.CaseService();
                service.Url = classService.URL;

                string[] caseNumbers = service.GetCaseNumber();

                RibbonDropDownItem item;
                comboBox1.Items.Clear();
                foreach (string caseNumber in caseNumbers)
                {
                    item = Factory.CreateRibbonDropDownItem();
                    item.Label = caseNumber.ToString();
                    comboBox1.Items.Add(item);
                }

                service.Dispose();
                service = null;
            }
            catch
            {
            }
        }
        #endregion
    }
}
