using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Diagnostics;
using Microsoft.Office.Core;
using Microsoft.Win32;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace TestOutlookAddIn
{
    class Processing
    {

        public Microsoft.Office.Interop.Outlook.Application toOtrs_2020;
        public Param Objparam;

        public Processing(Microsoft.Office.Interop.Outlook.Application _this, Param _param)
        {
            toOtrs_2020 = _this;
            Objparam = _param;

        }
        public void copyMailItem()
        {

            //MessageBox.Show("Juhu");

            //            Outlook.MailItem mail = this.Application.GetNamespace("MAPI").
            Outlook.MailItem mail = toOtrs_2020.Application.GetNamespace("MAPI").
                  GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).
      Items.GetFirst() as Outlook.MailItem;


            //Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)this.Application.
            Outlook.MAPIFolder inBox = (Outlook.MAPIFolder)toOtrs_2020.Application.
         ActiveExplorer().Session.GetDefaultFolder
         (Outlook.OlDefaultFolders.olFolderInbox);
            // Outlook.Items items = (Outlook.Items)inBox.Items;
            //Object selObject = this.Application.ActiveExplorer().Selection[1];
            Object selObject = toOtrs_2020.Application.ActiveExplorer().Selection[1];
            Outlook.MailItem selectedMail = (selObject as Outlook.MailItem);

            //Outlook.MailItem moveMail = null;
            //items.Restrict("[UnRead] = true");

            //Outlook.MAPIFolder destFolder = inBox.Folders["Test"];

            // Dialog zum TicketID erfassen anzeigen.
            Objparam.Subject = selectedMail.Subject;
            getTicketID _getTicketID = new getTicketID(Objparam);

            _getTicketID.ShowDialog();

            // Outlook.MAPIFolder destFolder = this.Application.Session.PickFolder() as Outlook.Folder;

            //Outlook.MAPIFolder destFolder = DestFolder();
            string _folderID = getDestFolder();

            if (_folderID != null)
            {
                //                Outlook.MAPIFolder destFolder = (Outlook.MAPIFolder)this.Application.ActiveExplorer().Session.GetFolderFromID(_folderID);
                Outlook.MAPIFolder destFolder = (Outlook.MAPIFolder)toOtrs_2020.Application.ActiveExplorer().Session.GetFolderFromID(_folderID);



                if (selectedMail != null && Objparam.IsChecked != false)
                {
                    // Create a copy of the item.
                    Outlook.MailItem copyMail = selectedMail.Copy() as Outlook.MailItem;

                    // Objparam.NewSubject wurde in getTicketID() ermittelt.
                    //copyMail.Subject = Objparam.NewSubject;

                    string MailTime = selectedMail.ReceivedTime.ToString();

                    copyMail.Subject = Objparam.Subject + " - " + MailTime;

                    // kopierte Mail verschieben.
                    try
                    {
                        copyMail.Move(destFolder);

                    }
                    catch
                    {
                        MessageBox.Show("Zielordner nicht gefunden. Bitte Registry Key löschen und Pfad neu setzen.\n HKEY_CURRENT_USER\\SOFTWARE\\BDG\\ToOTRS");
                    }

                }

            }
        }



        private string getDestFolder()
        {
            //Outlook.MAPIFolder _folder = null;

            string _folderID = (string)Registry.GetValue(@Objparam.RegPath, Objparam.RegKey, null);

            if (_folderID != null)
            {
                return _folderID;
            }
            else
            {

                return setDestFolder();
            }
        }

        private string setDestFolder()
        {
            // Pfad auswählen und in Registry schreiben.
            MessageBox.Show("Keine Zielordner für OTRS Mails definiert. Bitte zuerst Zielordner auswählen!");

            //Outlook.MAPIFolder _folder = this.Application.Session.PickFolder() as Outlook.Folder;
            Outlook.MAPIFolder _folder = toOtrs_2020.Application.Session.PickFolder() as Outlook.Folder;


            if (_folder == null)
            {
                return null;
            }
            else
            {

                try
                {
                    Registry.SetValue(@Objparam.RegPath, Objparam.RegKey, _folder.EntryID);
                    return _folder.EntryID;
                }
                catch
                {
                    MessageBox.Show("Fehler beim schreiben in die Registry.");
                    return null;
                }
            }

        }

    }
}
