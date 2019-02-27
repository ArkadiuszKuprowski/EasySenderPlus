using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using System.DirectoryServices.AccountManagement;
using System.Windows.Forms;
using System.ComponentModel;
using System.Security.Principal;
using Microsoft.Office.Interop.Outlook;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Exchange.WebServices.Data;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using easyDMSTool.Properties;
using Application = Microsoft.Office.Interop.Outlook.Application;
using Exception = System.Exception;
using System.Xml;
using System.Drawing;


//using Microsoft.Exchange.WebServices.Data;
// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new ribbonEasyDMS();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace easyDMSTool
{
    [ComVisible(true)]
    public class ribbonEasyDMS : Office.IRibbonExtensibility
    {

        #region Properties
        private Office.IRibbonUI ribbon;
        private string dateFormat = "yyyyMMddHHmmssffff";
        private static string serverUrl = Settings.Default.serverUrl;
        private static string userID = Settings.Default.userID;
        private static string userPassword = Settings.Default.userPassword;
        public static List<mappingAD> mapAD;
        public static List<string> groupAD;
        public static HashSet<string> groupAD_hash;
        public static List<string> currentUserGroupList;
        public static HashSet<string> currentUserGroupList_hash;
        public static string fileLog = null;
        public static FileStream myTraceLog;
        public static StreamWriter file;
        #endregion

        public ribbonEasyDMS()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetRemoteResourceText("\\\\EDMS\\EasySender\\config\\DocumentTypes.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            initializeMembers();
        }
        public System.Drawing.Bitmap getCustomImage(Office.IRibbonControl control)
        {
           string name = String.Concat("\\\\EDMS\\EasySender\\config\\flags\\", control.Id.TrimEnd(new char[] { '_','n', 't', 'b' }) + ".png");
           //return (System.Drawing.Bitmap)Resources.ResourceManager.GetObject(name);
           Bitmap pic = new Bitmap(name);
           return pic;
        }

        public Bitmap getRemoteImg(String fileName ){
            String completeName = Path.Combine(@"C:\Users\Public\ES\", fileName);
            Bitmap pic = new Bitmap(completeName);
            return pic;
        }

        public void getCurrentStatus(Office.IRibbonControl control)
        {
            System.Windows.Forms.MessageBox.Show("Server: " + ribbonEasyDMS.serverUrl + "\nUser ID: " + ribbonEasyDMS.userID);
            //TODO: add build version
        }
        
        public bool isCountryVisible(Office.IRibbonControl control) =>
               currentUserGroupList.Exists(entry => entry.Contains("-" + control.Tag + "-"));

        public Boolean isDoctypeVisible(Office.IRibbonControl control)
        {
            string[] strArray = control.Tag.Split(new char[] { ';' });
            string searchablePattern = "-" + strArray[1] + "-";
            searchablePattern = (strArray[3] != "Finance") ? (searchablePattern + "CS") : (searchablePattern + strArray[3]);
            return currentUserGroupList.Exists(x => x.Contains(searchablePattern));
        }

        public void onButtonClicked(Office.IRibbonControl control)
        {
            string[] strArray = control.Tag.Split(new char[] { ';' });
            this.ExtractEmail(strArray[2], strArray[3], strArray[0], strArray[1], control.Id);
        }

        #endregion

        #region Helpers

        private static string GetRemoteResourceText(string resourceName)
        {
            string str="";
            StreamReader reader = new StreamReader(resourceName);
            str = reader.ReadToEnd();
            reader.Close();                              
            return str;
        }

        public void initializeMembers() //TODO: clean up objects initialization
        {
            if (!Directory.Exists(@"C:\Users\Public\pdfWork"))
            {
                Directory.CreateDirectory(@"C:\Users\Public\pdfWork");
            }
            Settings.Default.userID = UserPrincipal.Current.ToString(); 
            currentUserGroupList = new List<string>();
            mapAD = mappingAD.Deserialize();
            groupAD = new List<string>();
            for (int i = 0; i < mapAD.Count; i++)
            {
                groupAD.Add(mapAD[i].groupAD);
            }
            foreach (Principal principal in UserPrincipal.Current.GetGroups())
            {
                string groupName = principal.ToString();
                if (groupAD.Exists(x => x.Contains(groupName)))
                {
                    currentUserGroupList.Add(groupName);
                }
            }
            groupAD_hash = new HashSet<string>(groupAD);
            currentUserGroupList.Sort();
            currentUserGroupList_hash = new HashSet<string>(currentUserGroupList);
        }


        private void ExtractEmail(string country, string fileType, string docType, string countryCode, string emailType)
        {
            Application application = Globals.ThisAddIn.Application;
            if (application.ActiveExplorer().Selection.Count > 0)
            {
                foreach (MailItem item in application.ActiveExplorer().Selection)
                {
                    string folder = this.GetOutputFolder(country, docType, item);
                    if (item.Attachments.Count <= 0)
                    {
                        if (MessageBox.Show("Do you really want to send this e-mail to EDMS?", "Sending e-mail without attachments.", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                        {
                            continue;
                        }
                        this.ExtractEmailBody(item, folder, "Message_Body.rtf");
                        this.ConvertFiles(folder, country, fileType, docType, countryCode, item);
                        continue;
                    }
                    List<string> list = new List<string>();
                    int num = 1;
                    while (true)
                    {
                        if (num > item.Attachments.Count)
                        {
                            SelectAttachments attachments = new SelectAttachments(list);
                            if (attachments.ShowDialog() == DialogResult.OK)
                            {
                                if ((attachments.selectedAttachments.Capacity == 0) && attachments.withMailBody)
                                {
                                    this.ExtractEmailBody(item, folder, "Message_Body.rtf");
                                    this.ConvertFiles(folder, country, fileType, docType, countryCode, item);
                                }
                                else if ((attachments.selectedAttachments.Capacity > 0) && !attachments.withMailBody)
                                {
                                    this.ExtractEmailAttachments(application, item, folder, attachments.selectedAttachments);
                                    this.ConvertFiles(folder, country, fileType, docType, countryCode, item, attachments.selectedAttachments, "e-mail from " + item.SenderName + " " + this.GetSenderSMTPAddress(item));
                                }
                                else if ((attachments.selectedAttachments.Capacity > 0) && attachments.withMailBody)
                                {
                                    this.ExtractEmailBody(item, folder, "Message_Body.rtf");
                                    this.ExtractEmailAttachments(application, item, folder, attachments.selectedAttachments);
                                    this.ConvertFiles(folder, country, fileType, docType, countryCode, item, attachments.selectedAttachments, "e-mail from " + item.SenderName + " " + this.GetSenderSMTPAddress(item));
                                }
                                else if ((attachments.selectedAttachments.Capacity == 0) && !attachments.withMailBody)
                                {
                                    MessageBox.Show("You did not select either attachment or e-mail body. \n\nThere is nothing to be sent. \n\n Please get back and make your choice. ", "Empty selection.", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                                }
                            }
                            break;
                        }
                        if (!item.Attachments[num].FileName.EndsWith(".eml") && !item.Attachments[num].FileName.EndsWith(".msg"))
                        {
                            list.Add(item.Attachments[num].FileName);
                        }
                        else
                        {
                            string path = Path.Combine(@"C:\Users\Public\pdfWork\", item.Attachments[num].FileName);
                            item.Attachments[num].SaveAsFile(path);
                            MailItem item2 = Globals.ThisAddIn.Application.Session.OpenSharedItem(path) as MailItem;
                            list.Add(item2.Subject.Replace(";", "") + "_Message_Body");
                            int num2 = 1;
                            while (true)
                            {
                                if (num2 > item2.Attachments.Count)
                                {
                                    try
                                    {
                                        File.Delete(path);
                                    }
                                    catch (SystemException exception1)
                                    {
                                        string message = exception1.Message;
                                    }
                                    break;
                                }
                                list.Add(item2.Subject + "_" + item2.Attachments[num2].FileName);
                                num2++;
                            }
                        }
                        num++;
                    }
                }
            }
        }

        private bool ExtractEmailAttachments(Application application, MailItem objMail, string folder, List<string> attachments)
        {
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            if (objMail.Attachments.Count > 0)
            {
                int count = objMail.Attachments.Count;
                int num2 = 1;
                while (true)
                {
                    if (num2 > count)
                    {
                        objMail.Save();
                        break;
                    }
                    if (!objMail.Attachments[num2].FileName.EndsWith(".eml") && !objMail.Attachments[num2].FileName.EndsWith(".msg"))
                    {
                        if (attachments.Contains(objMail.Attachments[num2].FileName))
                        {
                            objMail.Attachments[num2].SaveAsFile(Path.Combine(folder, objMail.Attachments[num2].FileName));
                        }
                    }
                    else
                    {
                        string path = Path.Combine(folder, num2 + "_" + objMail.Attachments[num2].FileName);
                        objMail.Attachments[num2].SaveAsFile(path);
                        MailItem item = application.Session.OpenSharedItem(path) as MailItem;
                        string str2 = item.Subject.Replace(":", "");
                        int num3 = item.Attachments.Count;
                        int num4 = 1;
                        while (true)
                        {
                            if (num4 > num3)
                            {
                                if (attachments.Contains(item.Subject + "_Message_Body"))
                                {
                                    this.ExtractEmailBody(item, folder, string.Concat(new object[] { num2, "_0_", str2, "_Message_Body.rtf" }));
                                }
                                break;
                            }
                            if (attachments.Contains(item.Subject + "_" + item.Attachments[num4].FileName))
                            {
                                object[] objArray = new object[] { num2, "_", num4, "_", str2, "_", item.Attachments[num4].FileName };
                                item.Attachments[num4].SaveAsFile(Path.Combine(folder, string.Concat(objArray)));
                            }
                            num4++;
                        }
                    }
                    num2++;
                }
            }
            return true;
        }

        private bool ExtractEmailBody(MailItem objMail, string folder, string name)
        {
            try
            {
                string body = objMail.Body;
                string path = Path.Combine(folder, name);
                Directory.CreateDirectory(folder);
                objMail.SaveAs(path, OlSaveAsType.olRTF);
                return true;
            }
            catch (Exception exception1)
            {
                MessageBox.Show("Sth happened while working on file " + folder + name + ". \n Error caught: " + exception1.Message);
                return false;
            }
        }

        private void ConvertFiles(string folder, string country, string fileType, string docType, string countryCode, MailItem objMail)
        {
            BackgroundWorker worker = new BackgroundWorker();
            bool isSent = false;
            worker.DoWork += (sender, e) =>
            {
                FileConverter fc = new FileConverter();
                isSent = new FileConverter().Convert(folder, "", fileType, country, docType, serverUrl, countryCode, "email from " + objMail.SenderName + " " + this.GetSenderSMTPAddress(objMail));
            };
            worker.RunWorkerCompleted += (sender, e) => this.setDocumentType(objMail, docType, country, isSent);
            worker.RunWorkerAsync();
        }


        private void ConvertFiles(string folder, string country, string fileType, string docType, string countryCode, MailItem objMail, List<string> selectedAttachments, string emailSender)
        {
            BackgroundWorker worker = new BackgroundWorker();
            bool isSent = false;
            worker.DoWork += (sender, e) => {
                FileConverter fc = new FileConverter();
                isSent = new FileConverter().Convert(folder, "", fileType, country, docType, serverUrl, countryCode, emailSender);
            };
            worker.RunWorkerCompleted += (sender, e) => this.setDocumentType(objMail, docType, country, isSent);
            worker.RunWorkerAsync();
        }


        private void detectDuplicateAttachments(MailItem objMail)
        {
            Attachments attachments = objMail.Attachments;
            int count = attachments.Count;
            for (int i = 1; i < count; i++)
            {
                string path = Path.Combine(@"C:\Users\Public\pdfWork\", i + "_" + objMail.Attachments[i].FileName);
                attachments[i].SaveAsFile(path);
                attachments.Add(path, Missing.Value, Missing.Value, Missing.Value);
                File.Delete(path);
            }
            for (int j = 1; j < count; j++)
            {
                attachments[1].Delete();
            }
        }

        private void detectDuplicateAttachments(MailItem objMail, List<string> names)
        {
            Attachments attachments = objMail.Attachments;
            int count = attachments.Count;
            for (int i = 1; i < count; i++)
            {
                if (names.Contains(attachments[i].DisplayName))
                {
                    string path = Path.Combine(@"C:\Users\Public\pdfWork\", i + "_" + objMail.Attachments[i].FileName);
                    attachments[i].SaveAsFile(path);
                    attachments.Add(path, Missing.Value, Missing.Value, Missing.Value);
                    File.Delete(path);
                }
            }
            for (int j = 1; j < count; j++)
            {
                attachments[1].Delete();
            }
        }


        private string GetOutputFolder(string country, string docType, MailItem objMail)
        {
            string path = Path.Combine(Path.Combine(Path.Combine(@"C:\Users\Public\pdfWork", country), docType.Replace("/", "-")), objMail.GetHashCode().ToString() + "_" + DateTime.Now.ToString(this.dateFormat));
            if (Directory.Exists(path))
            {
                path = path + "_" + DateTime.Now.ToString(this.dateFormat);
            }
            return path;
        }


        public static string getUserID() => userID;

        private string GetSenderSMTPAddress(MailItem mail)
        {
            string schemaName = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
            if (mail == null)
            {
                throw new ArgumentNullException();
            }
            if (mail.SenderEmailType != "EX")
            {
                return mail.SenderEmailAddress;
            }
            AddressEntry sender = mail.Sender;
            if (sender == null)
            {
                return null;
            }
            if ((sender.AddressEntryUserType != OlAddressEntryUserType.olExchangeUserAddressEntry) && (sender.AddressEntryUserType != OlAddressEntryUserType.olExchangeRemoteUserAddressEntry))
            {
                return (sender.PropertyAccessor.GetProperty(schemaName) as string);
            }
            ExchangeUser exchangeUser = sender.GetExchangeUser();
            return exchangeUser?.PrimarySmtpAddress;
        }

           
        private void setDocumentType(MailItem objMail, string itemType, string country, bool isSent)
        {
            try
            {
                this.setDocumentType(objMail, country);
                if (isSent || (objMail == null))
                {
                    objMail.UserProperties["Document Type"].Value = itemType;
                }
                else
                {
                    itemType = "NOT DELIVERED";
                    objMail.UserProperties["Document Type"].Value = itemType;
                }

                objMail.Save();
            }
            catch (Exception e)
            {
                String caption = "OOPS";
                String message = "Snap! \n\nWhy have you deleted the email right after sending it to EDMS? \nPlease check it in Easy. If something got lost - recover email. Re-send. Check Easy.\n\nAnd here goes error call stack (share it with SD):\n\n" + e.StackTrace;
                DialogResult result;
                result = MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Hand);               
            }
        }

        private void setDocumentType(MailItem objMail, string country)
        {
            objMail.UserProperties.Add("TransferredBy", OlUserPropertyType.olText, true, OlUserPropertyType.olText);
            objMail.UserProperties.Add("Document Type", OlUserPropertyType.olText, true, OlUserPropertyType.olText);
            objMail.UserProperties.Add("Country", OlUserPropertyType.olText, true, OlUserPropertyType.olText);
            objMail.UserProperties["TransferredBy"].Value = WindowsIdentity.GetCurrent().Name;
            objMail.UserProperties["Country"].Value = country;
        }


        public static void setUserPassword(string pwd)
        {
            userPassword = pwd;
        }

        public static void setUserID(string id)
        {
            userID = id;
        }

        public static void setServerUrl(string url)
        {
            serverUrl = url;
        }

        #endregion
    }
}
