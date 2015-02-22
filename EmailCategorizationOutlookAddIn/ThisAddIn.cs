using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;

namespace EmailCategorizationOutlookAddIn
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;
        const string cVeryImportantEmail = "Very Important";
        const string cImportantEmail = "Important";

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            if (CreateRequiredFolders() == false) return;

            Outlook.MAPIFolder inbox = this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            inbox.Items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAdd);
            var unreadItems = inbox.Items.Restrict("[Unread]=true").Cast<Outlook.MailItem>().ToList();
            foreach (var item in unreadItems)
            {
                CategorizeEmail(item);
            }
        }

        void Items_ItemAdd(object Item)
        {
            Outlook.MailItem mail = (Outlook.MailItem)Item;
            CategorizeEmail(mail);
        }

        private void CategorizeEmail(Outlook.MailItem item)
        {
            if (item == null) return;
            try
            {
                Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)this.Application.ActiveExplorer().Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                var isVeryImportant = item.Body.IndexOf("Big Boss") > 0;
                if (isVeryImportant)
                    item.Move(inBox.Folders[cVeryImportantEmail]);
                var isImportant = item.Body.IndexOf("Small Boss") > 0;
                if (isImportant)
                    item.Move(inBox.Folders[cImportantEmail]);
            }
            catch (Exception ex)
            {
                MessageBox.Show("The following error occurred: " + ex.Message);
            }
        }

        private bool CreateRequiredFolders()
        {
            var requiredFolders = new string[] { cVeryImportantEmail, cImportantEmail };
            try
            {
                Outlook.MAPIFolder inBox = Application.Session.DefaultStore.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                var existingFolder = inBox.Folders.Cast<Outlook.MAPIFolder>().Where(r => requiredFolders.Contains(r.Name)).Select(r => r.Name).ToList();
                var createFolders = requiredFolders.Where(r => !existingFolder.Contains(r));
                foreach (var needFolder in createFolders)
                {
                    inBox.Folders.Add(needFolder, Outlook.OlDefaultFolders.olFolderInbox);
                }
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("The following error occurred: " + ex.Message);
                return false;
            }
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
