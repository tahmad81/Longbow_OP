using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Services;
using System.Web.Security;
using System.IO;
using System.Text;
using iwantedue;

/// <summary>
/// Summary description for CaseService
/// </summary>
[WebService(Namespace = "CaseService")]
[WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
// To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
// [System.Web.Script.Services.ScriptService]
public class CaseService : System.Web.Services.WebService
{

    public CaseService()
    {

        //Uncomment the following line if using designed components 
        //InitializeComponent(); 
    }

    [WebMethod]
    public string[] GetCaseNumber()
    {
        List<string> caseNumber = new List<string>();
        DataSet1TableAdapters.CasesTableAdapter caseTableAdapter = new DataSet1TableAdapters.CasesTableAdapter();
        foreach (DataSet1.CasesRow caseRow in caseTableAdapter.GetDataByNull().Rows)
        {
            string caseName = "";
            try
            {
                caseName = caseRow.Case_Name;
            }
            catch
            {
                caseName = "";
            }

            if (caseName != "")
            {
                caseNumber.Add(caseName + " : " + caseRow.File_Number);
            }
            else
            {
                caseNumber.Add(caseRow.File_Number);
            }


        }
        /*
        DataSet1TableAdapters.EmailsTableAdapter emailTableAdapter = new DataSet1TableAdapters.EmailsTableAdapter();
        foreach (DataSet1.EmailsRow emailRow in emailTableAdapter.GetDataByCaseID().Rows)
        {
            caseNumber.Add(emailRow.Case_Id);
        }
        */
        return caseNumber.ToArray();
    }

    [WebMethod]
    public string Uploads(MessageClass message, string user)
    {
        // Prepare storage folder
        string TEMP_PATH = Server.MapPath("temporary");

        if (!TEMP_PATH.EndsWith("\\"))
        {
            TEMP_PATH += "\\";
        }

        this.CheckFolder(TEMP_PATH);
        this.CheckFolder(TEMP_PATH + "_Attachment");

        Stream stream = new MemoryStream(message.Message);
        OutlookStorage.Message storage = new OutlookStorage.Message(stream);
        if (storage.Attachments.Count > 0)
        {
            message.Attachment = true;
        }
        else
        {
            message.Attachment = false;
        }

        DataSet1TableAdapters.CasesTableAdapter caseAdapter = new DataSet1TableAdapters.CasesTableAdapter();

        string caseID = message.CaseID;

        if (caseID.Contains(":"))
        {
            caseID = caseID.Substring(caseID.IndexOf(": ") + 2);
        }


        foreach (DataSet1.CasesRow row in caseAdapter.GetDataByFileNumber(caseID).Rows)
        {
            caseID = row.Case_Id.ToString();
        }

        // Parse message here
        //DataSet1TableAdapters.EmailsTableAdapter emailAdapter = new DataSet1TableAdapters.EmailsTableAdapter();

        //emailAdapter.InsertQuery(Convert.ToInt32(caseID),
        //                                message.EntryID,
        //                                message.Sender,
        //                                message.Recipient,
        //                                message.DateSent,
        //                                message.DateReceived,
        //                                message.Subject,
        //                                message.Body,
        //                                message.Message,
        //                                message.Attachment);

        long mailID = 0;
        using (Longbow_TauseefModel.Longbow_TauseefEntities context = new Longbow_TauseefModel.Longbow_TauseefEntities())
        {
            Longbow_TauseefModel.Email emailToSave = new Longbow_TauseefModel.Email()
            {
                Case_Id = int.Parse(caseID),
                Entry_Id = message.EntryID,
                Sender = message.Sender,
                Recipient = message.Recipient,
                Date_Sent = message.DateSent,
                Date_Recieved = message.DateReceived,
                Subject = message.Subject,
                Body = message.Body,
                Attachments = message.Attachment,
                Message_Object = message.Message
            };
            context.AddToEmails(emailToSave);
            context.SaveChanges();
            mailID = emailToSave.Mail_Id;
        }
        //foreach (DataSet1.EmailsRow emailRow in emailAdapter.GetDataByEntryID(message.EntryID))
        //{
        //    mailID = emailRow.Mail_Id;
        //}

       // emailAdapter = null;

        if (mailID != 0)
        {
            DataSet1TableAdapters.Email_AttachmentsTableAdapter attachmentAdapter = new DataSet1TableAdapters.Email_AttachmentsTableAdapter();
            //Stream stream = new MemoryStream(message.Message);
            //OutlookStorage.Message storage = new OutlookStorage.Message(stream);

            foreach (OutlookStorage.Attachment attachment in storage.Attachments)
            {
                string fileName = attachment.Filename;
                byte[] contents = attachment.Data;
                attachmentAdapter.InsertQuery(mailID, fileName, contents);
            }
            storage.Dispose();

            attachmentAdapter = null;

            // Fill the relationship
            // --- Get user ID
            DataSet1TableAdapters.UsersTableAdapter userTableAdapter = new DataSet1TableAdapters.UsersTableAdapter();

            int userID = 0;
            foreach (DataSet1.UsersRow userRow in userTableAdapter.GetDataByUsername(user))
            {
                userID = userRow.User_ID;
                break;
            }
            userTableAdapter = null;

            // --- Insert Relations//comment for now
            //DataSet1TableAdapters.RelationsTableAdapter relationAdapter = new DataSet1TableAdapters.RelationsTableAdapter();
            //relationAdapter.InsertQuery(mailID, userID);
            //relationAdapter = null;
        }



        return Context.User.Identity.Name;
    }

    [WebMethod]
    public bool Login(string strUser, string strPwd)
    {
        bool strRole = AuthenticateUser(strUser, strPwd);

        if (strRole)
        {
            FormsAuthentication.SetAuthCookie(strUser, false);
            return true;
        }
        else
            return false;
    }

    [WebMethod]
    public void LogOut()
    {
        FormsAuthentication.SignOut();
    }

    [WebMethod]
    public bool IsValidated()
    {
        if (Context.User.Identity.Name != null && Context.User.Identity.IsAuthenticated)
        {
            return false;
        }
        else
        {
            FormsAuthentication.SignOut();
            return false;
        }
    }

    #region "Private Methods"
    private bool AuthenticateUser(string username, string password)
    {
        if (username == "" || password == "")
            return false;

        DataSet1TableAdapters.UsersTableAdapter userTableAdapter = new DataSet1TableAdapters.UsersTableAdapter();
        try
        {
            int count = (int)userTableAdapter.ScalarQuery(username, password);
            if (count == 0)
            {
                return false;
            }

        }
        catch
        {
            return false;
        }

        return true;
    }

    private string RandomString(int size, bool lowerCase)
    {
        StringBuilder builder = new StringBuilder();
        Random random = new Random();
        char ch;
        for (int i = 0; i < size; i++)
        {
            ch = Convert.ToChar(Convert.ToInt32(Math.Floor(26 * random.NextDouble() + 65)));
            builder.Append(ch);
        }
        if (lowerCase)
            return builder.ToString().ToLower();
        return builder.ToString();
    }

    private void CheckFolder(string folderName)
    {
        if (!Directory.Exists(folderName))
        {
            Directory.CreateDirectory(folderName);
        }

    }

    private string StoreMessage(string path, byte[] data)
    {
        string fileName = path + RandomString(10, true) + ".eml";
        try
        {
            FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Write);
            fs.Write(data, 0, data.Length);
            fs.Close();

            fs = null;
        }
        catch (Exception ex)
        {
            return ex.Message.ToString();
        }

        return "";
    }
    #endregion

}

public class MessageClass
{
    public MessageClass()
    {
    }

    public MessageClass(string caseID, string entryID, string sender, string recipient, DateTime dateSent, DateTime dateReceived, string subject, string body, bool attachment, byte[] message)
    {
        CaseID = caseID;
        EntryID = entryID;
        Sender = sender;
        Recipient = recipient;
        DateReceived = dateReceived;
        DateSent = dateSent;
        Subject = subject;
        Body = body;
        Message = message;
        Attachment = attachment;
    }

    public string CaseID { get; set; }
    public string EntryID { get; set; }
    public string Sender { get; set; }
    public string Recipient { get; set; }
    public DateTime DateReceived { get; set; }
    public DateTime DateSent { get; set; }
    public string Subject { get; set; }
    public string Body { get; set; }
    public bool Attachment { get; set; }
    public byte[] Message { get; set; }

}
