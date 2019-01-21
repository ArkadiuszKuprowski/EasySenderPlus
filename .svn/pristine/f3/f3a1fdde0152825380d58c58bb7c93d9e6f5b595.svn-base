using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace easyDMSTool
{
    [System.Runtime.InteropServices.ComVisible(true)]
    public partial class easyDMSToolOptionDialog : UserControl, Microsoft.Office.Interop.Outlook.PropertyPage
    {
        public easyDMSToolOptionDialog()
        {            
            InitializeComponent();
            this.Load += new EventHandler(easyDMSToolOptionDialog_Load);
        }

        void Microsoft.Office.Interop.Outlook.PropertyPage.Apply()
        {
            if (isDirty)
            {               
                if (serverOther_rbtn.Checked) 
                    saveServerName();

                if (useProvidedUser_rbtn.Checked) 
                    saveUserCredentials();

                easyDMSTool.Properties.Settings.Default.Save();
            }
        }

        bool Microsoft.Office.Interop.Outlook.PropertyPage.Dirty
        {
            get
            {
                return isDirty;
            }
        }

        void Microsoft.Office.Interop.Outlook.PropertyPage.GetPageInfo(ref string helpFile, ref int helpContext)
        {
            
        }

        [System.Runtime.InteropServices.DispId(captionDispID)]
        public string PageCaption
        {
            get
            {
                return "Send To EasyDMS settings";
            }
        }

        private void easyDMSToolOptionDialog_Load(object sender, System.EventArgs e)
        {
            Type myType = typeof(System.Object);
            string assembly =
            System.Text.RegularExpressions.Regex.Replace(myType.Assembly.CodeBase, "mscorlib.dll", "System.Windows.Forms.dll");
            assembly = System.Text.RegularExpressions.Regex.Replace(assembly, "file:///", "");
            assembly = System.Reflection.AssemblyName.GetAssemblyName(assembly).FullName;
            Type unmanaged =
                Type.GetType(System.Reflection.Assembly.CreateQualifiedName
                (assembly, "System.Windows.Forms.UnsafeNativeMethods"));
            Type oleObj = unmanaged.GetNestedType("IOleObject");
            System.Reflection.MethodInfo mi = oleObj.GetMethod("GetClientSite");
            object myppSite = mi.Invoke(this, null);
            this.ppSite = (Microsoft.Office.Interop.Outlook.PropertyPageSite)myppSite;
        }

        private void onDirty()
        {
            isDirty = true;
            ppSite.OnStatusChange();
        }


        #region button methods

        private void serverTest_rbtn_CheckedChanged(object sender, EventArgs e)
        {
            easyDMSTool.Properties.Settings.Default.serverUrl = "deis366";
            easyDMSTool.Properties.Settings.Default.isCheckedServerTest_rbtn = true;
            easyDMSTool.Properties.Settings.Default.isCheckedServerOther_rbtn = false;
            easyDMSTool.Properties.Settings.Default.isCheckedServerProd_rbtn = false;
            ribbonEasyDMS.setServerUrl("deis366");
            easyDMSTool.Properties.Settings.Default.Save();
            onDirty();          
        }
        private void serverProd_rbtn_CheckedChanged(object sender, EventArgs e)
        {

            easyDMSTool.Properties.Settings.Default.serverUrl = "deis335";
            easyDMSTool.Properties.Settings.Default.isCheckedServerProd_rbtn = true;
            easyDMSTool.Properties.Settings.Default.isCheckedServerTest_rbtn = false;
            easyDMSTool.Properties.Settings.Default.isCheckedServerOther_rbtn = false;
            ribbonEasyDMS.setServerUrl("deis335");
            onDirty();
        }
        private void serverOther_rbtn_CheckedChanged(object sender, EventArgs e)
        {
            if (serverOther_rbtn.Checked)
            {
                serverOther_txtbox.Enabled = true;
                easyDMSTool.Properties.Settings.Default.isCheckedServerOther_rbtn = true;
                easyDMSTool.Properties.Settings.Default.isCheckedServerTest_rbtn = false;                
                easyDMSTool.Properties.Settings.Default.isCheckedServerProd_rbtn = false;
                onDirty();
            }
            else
            {
                serverOther_txtbox.Enabled = false;
                easyDMSTool.Properties.Settings.Default.isCheckedServerOther_rbtn = false;
                onDirty();
            }
        }

        private void saveServerName()
        {
            if (serverOther_rbtn.Checked)
            {
                if (serverOther_txtbox.Text != "")
                {
                    easyDMSTool.Properties.Settings.Default.serverUrlOther = serverOther_txtbox.Text;
                    easyDMSTool.Properties.Settings.Default.serverUrl = serverOther_txtbox.Text;
                    ribbonEasyDMS.setServerUrl(serverOther_txtbox.Text);

                    onDirty();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Check if server name is not empty! ");
                }
            }
        }
        private void saveUserCredentials()
        {
            if (userID_txtbox.Enabled)
            {
                if (userID_txtbox.Text != "" )
                {
                    easyDMSTool.Properties.Settings.Default.userID = userID_txtbox.Text;
                    ribbonEasyDMS.setUserID(easyDMSTool.Properties.Settings.Default.userID);
                    ribbonEasyDMS.setUserPassword(easyDMSTool.Properties.Settings.Default.userPassword);
                    onDirty();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Check if textboxes are not empty");
                }
            }
        }

        private void useDefaultUser_rbtn_CheckedChanged(object sender, EventArgs e)
        {
                userID_txtbox.Enabled = false;
                ribbonEasyDMS.setUserID(easyDMSTool.Properties.Settings.Default.userIDDefault);
                ribbonEasyDMS.setUserPassword(easyDMSTool.Properties.Settings.Default.userPasswordDefault);
                easyDMSTool.Properties.Settings.Default.userID = easyDMSTool.Properties.Settings.Default.userIDDefault;
                easyDMSTool.Properties.Settings.Default.userPassword = easyDMSTool.Properties.Settings.Default.userPasswordDefault;
                easyDMSTool.Properties.Settings.Default.isCheckedUseDefaultUser_rbtn = true;
                easyDMSTool.Properties.Settings.Default.isCheckedUseProvidedUser_rbtn = false;
                easyDMSTool.Properties.Settings.Default.isEnabledUseProvidedUser_txtbox = false;  
                onDirty();
        }
        private void useProvidedUser_rbtn_CheckedChanged(object sender, EventArgs e)
        {
                userID_txtbox.Enabled = true;
                easyDMSTool.Properties.Settings.Default.isCheckedUseProvidedUser_rbtn = true;
                easyDMSTool.Properties.Settings.Default.isCheckedUseDefaultUser_rbtn = false;
                easyDMSTool.Properties.Settings.Default.isEnabledUseProvidedUser_txtbox = true;
                onDirty();
        }

        #endregion


        const int captionDispID = -518;
        private Microsoft.Office.Interop.Outlook.PropertyPageSite ppSite;
        
        bool isDirty = false;

        private void userID_txtbox_TextChanged(object sender, EventArgs e)
        {
            onDirty();
        }

        private void password_txtbox_TextChanged(object sender, EventArgs e)
        {
            onDirty();
        }
    }
}
