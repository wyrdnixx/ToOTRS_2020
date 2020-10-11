using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using TestOutlookAddIn.Properties;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;


// TODO:  Führen Sie diese Schritte aus, um das Element auf dem Menüband (XML) zu aktivieren:

// 1: Kopieren Sie folgenden Codeblock in die ThisAddin-, ThisWorkbook- oder ThisDocument-Klasse.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Erstellen Sie Rückrufmethoden im Abschnitt "Menübandrückrufe" dieser Klasse, um Benutzeraktionen
//    zu behandeln, z.B. das Klicken auf eine Schaltfläche. Hinweis: Wenn Sie dieses Menüband aus dem Menüband-Designer exportiert haben,
//    verschieben Sie den Code aus den Ereignishandlern in die Rückrufmethoden, und ändern Sie den Code für die Verwendung mit dem
//    Programmmodell für die Menübanderweiterung (RibbonX).

// 3. Weisen Sie den Steuerelementtags in der Menüband-XML-Datei Attribute zu, um die entsprechenden Rückrufmethoden im Code anzugeben.  

// Weitere Informationen erhalten Sie in der Menüband-XML-Dokumentation in der Hilfe zu Visual Studio-Tools für Office.


namespace TestOutlookAddIn
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {

        private Office.IRibbonUI ribbon;

        private ToOTRS_2020 app;

        public Ribbon1(ToOTRS_2020 _app)
        {
            app = _app;
        }


        public string GetCustomUI(string RibbonID)
        {
            return
    @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
    <contextMenus>    
        <contextMenu idMso=""ContextMenuMailItem"">
            <button 
                id=""MyContextMenuMailItem""
                label=""ToOTRS""               
                getImage=""GetIcon""
                showImage = ""true""
                onAction =""RibbonMenuClick""
                />
        </contextMenu>  
    </contextMenus>
</customUI>
";
        }


        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;

            
        }

        public Image GetIcon(Office.IRibbonControl control)
        {
            return Resources.icon;
        }

        public  void RibbonMenuClick(IRibbonControl control)
        {

            /* TESTS
             * var selection = control.Context as Microsoft.Office.Interop.Outlook.Selection;
            var mailItems = selection.OfType<Microsoft.Office.Interop.Outlook.MailItem>().ToList();
            MessageBox.Show(mailItems[0].Subject);
            mailItems[0].UnRead = true;
            */


            /*
            Outlook.Selection sel = control.Context as Outlook.Selection;
            Outlook.MailItem mail = sel[1];
            MessageBox.Show(mail.Body);

            */


            ToOTRS_2020.RightClickCallback(control,  app);


        }





        #region Hilfsprogramme

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }

}
