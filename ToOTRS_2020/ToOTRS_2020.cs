using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;

namespace TestOutlookAddIn
{
    public partial class ToOTRS_2020
    {

        //Outlook.Inspectors inspectors;
        //Outlook.Application thisApp;

        Param Objparam = new Param();
        private Processing processing;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            
            /*
             * inspectors = this.Application.Inspectors;
            inspectors.NewInspector +=
            new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
            */

            //this.Application.ItemContextMenuDisplay += new Outlook.ApplicationEvents_11_ItemContextMenuDisplayEventHandler(Application_ItemContextMenuDisplay2);

            Objparam.RegPath = "HKEY_CURRENT_USER\\SOFTWARE\\BDG\\ToOTRS";
            Objparam.RegKey = "DestFolder";
            // Versionsinfo auslesen und an Parameterobjekt übergeben.
            Assembly assembly = Assembly.GetExecutingAssembly();
            FileVersionInfo fileVersionInfo = FileVersionInfo.GetVersionInfo(assembly.Location);
            string version = fileVersionInfo.ProductVersion;

            Objparam.Version = version;

            processing = new Processing(this.Application, Objparam);

        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Hinweis: Outlook löst dieses Ereignis nicht mehr aus. Wenn Code vorhanden ist, der 
            //    muss ausgeführt werden, wenn Outlook heruntergefahren wird. Weitere Informationen finden Sie unter https://go.microsoft.com/fwlink/?LinkId=506785.
        }

      
        /*
        
        private void Application_ItemContextMenuDisplay2(Office.CommandBar CommandBar, Outlook.Selection Selection)
        {

            if (Selection[1] is Outlook.MailItem)
            {

                Outlook.MailItem selectedMailItem = Selection[1] as Outlook.MailItem;

               // this.CustomContextMenu(CommandBar, Selection);

            }

        }

        
        
        /// <summary>
        /// Outlook AddIn Kontext Menü
        /// </summary>
        /// <param name="CommandBar"></param>
        /// <param name="Selection"></param>
        private void CustomContextMenu(Office.CommandBar CommandBar, Outlook.Selection Selection)
        {

            Office.CommandBarButton customContextMenuTag = (Office.CommandBarButton)

                                                                                                CommandBar.Controls.Add

            (Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);

            customContextMenuTag.Click += new

            Office._CommandBarButtonEvents_ClickEventHandler(customContextMenuTag_Click);

            //customContextMenuTag.Caption = "Send to OTRS";

            // ?!?!?
            customContextMenuTag.FaceId = 351; //displays the image for the menu item
                                               //customContextMenuTag.FaceId = 1; //displays the image for the menu item

            customContextMenuTag.Style = Microsoft.Office.Core.MsoButtonStyle.msoButtonIconAndCaption;

        }
        
        */
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1(this);
        }


        public static void RightClickCallback(IRibbonControl control, TestOutlookAddIn.ToOTRS_2020 app)
        {
            Outlook.Selection sel = control.Context as Outlook.Selection;
            Outlook.MailItem mail = sel[1];
            // MessageBox.Show(mail.Body);
            //   copyMailItem();


            app.processing.copyMailItem();

            //processing.copyMailItem();
        }


        /*
        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    mailItem.Subject = "This text was added by using code";
                    mailItem.Body = "This text was added by using code";
                }

            }
        }
        */

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
