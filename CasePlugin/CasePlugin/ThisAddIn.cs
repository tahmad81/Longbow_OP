using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Tools.Ribbon;

namespace CasePlugin
{
    public partial class ThisAddIn
    {
        public bool useCaseNumber = false;
        public int intCaseNumber = 0;

        Outlook.Items items;

        ClassService classService = new ClassService();
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            useCaseNumber = false;
            items = Application.GetNamespace("MAPI").GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).Items;
            items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);
        }

        void items_ItemAdd(object Item)
        {
            if (!useCaseNumber)
                return;

            Outlook.MailItem mailItem;
            try
            {
                mailItem = (Outlook.MailItem)Item;
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

            messageClass.CaseID = intCaseNumber;
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

            // Load as byte array, then send
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
            if (caseService.Login(classService.Username, classService.Password))
            {
                MessageBox.Show("Invalid credential. Please check on configuration");
                return;
            }

            caseService.Uploads(messageClass, classService.Username);
            caseService.Dispose();
            caseService = null;


            try
            {
                File.Delete(classService.MSGPath + fileName);
            }
            catch
            { }

            
            //throw new NotImplementedException();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
