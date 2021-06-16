using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Microsoft.Win32;
using System.DirectoryServices;


namespace TestOutlookAddIn
{
    public partial class getTicketID : Form
    {
        //string TicketID = null;
        private Param _param = null;

        public getTicketID(Param param)
        {
            InitializeComponent();

            _param = param;
            
            // Überprüfungsstatus initial auf False setzen.
            _param.IsChecked = false;
            
            txtbox_Subject.Text = _param.Subject;
            // Textfeld Info leeren.
            lbl_info.Text = "";

            // Versionsinfo anzeigen
            label_info.Text = "ToOTRS © 2020 JoHe - v" + param.Version;

        }



        private void processForm()
        {           
            

            _param.Subject = txtbox_Subject.Text;


            if (checkNewTicket.Checked == true)
            {
                _param.IsChecked = true;
            } else
            {
                _param.IsChecked = processTicketNumber();
            }
            

            
            if (_param.IsChecked == true)
            {
                string surname = getADSurname();

                // MessageBox.Show("Surname: " + surname);

                if (surname == null)
                {
                    MessageBox.Show("Fehler beim ermitteln des Usernamens.");

                }
                else
                {
                    _param.Subject = _param.Subject + " # ToOTRS: " + surname;
                    this.Close();
                }                          
            }
            
        }



        private bool processTicketNumber()
        {
            //string tmp = tbox_TicketID.Text;

            // Emthällt der Betreff eine [MCB#<TICKETNUMMER>] ?
            Regex TicketIDinSubject = new Regex(@"(\[MCB\#[0-9]{16}\])");

            if (TicketIDinSubject.IsMatch(txtbox_Subject.Text))
            {
                _param.Subject = txtbox_Subject.Text;
                return true;
                
            }
            else  // überprüfe Feld TicketID
            {

                // Leerzeichen am Anfang oder Ende entfernen.
                tbox_TicketID.Text = tbox_TicketID.Text.Trim();

                // Überprüfen ob es eine gültige Ticketnummer ist (länge und nur Zahlen).
                Regex rgx = new Regex(@"^([0-9]{16})$");
                if (rgx.IsMatch(tbox_TicketID.Text))
                {
                    _param.TicketID = tbox_TicketID.Text;

                    _param.Subject = "[MCB#" + tbox_TicketID.Text + "]" + txtbox_Subject.Text;

                    return true;
                 
                    //return true;
                }
                else
                {
                    lbl_info.Text = DateTime.UtcNow + "Ticketnummer ist ungülltig (RegEx).";
                    return false;
                }
            }

        }

        private void btn_resetdstfolder_Click(object sender, EventArgs e)
        {

          //  string keyName = "\SOFTWARE\BDG\ToOTRS";
            
            try {
                Registry.CurrentUser.DeleteSubKeyTree("SOFTWARE\\BDG\\ToOTRS");
                lbl_info.Text = DateTime.UtcNow + ": Registrykey gelöscht!";
            } catch {
                lbl_info.Text = DateTime.UtcNow + ": Registrykey nicht vorhanden.";
            }
        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            processForm();  
        }



        private void tbox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                processForm();
            }

        }


        public string getADSurname()
        {

            string ADsurname;
            string curentUserName = Environment.UserName;

            try
            {

                DirectoryEntry entry1 = new DirectoryEntry("LDAP://RootDSE");
                string domainContext = entry1.Properties["defaultNamingContext"].Value as string;
                // Use the default naming context as the connected context may not work for searches
                DirectoryEntry entry = new DirectoryEntry("LDAP://" + domainContext);
                DirectorySearcher adSearch = new DirectorySearcher(entry);

                adSearch.Filter = "(&(objectClass=user)(anr=" + curentUserName + "))";

                foreach (SearchResult sResultSet in adSearch.FindAll())
                {
                    // Last Name
                    ADsurname = GetProperty(sResultSet, "sn");
                    lbl_info.Text = DateTime.UtcNow + ": AD Surname = " + ADsurname;

                    if (ADsurname == "")
                    {
                        lbl_info.Text = DateTime.UtcNow + ": AD Surname ist leer - Verwende Windows Username: " + curentUserName;

                        MessageBox.Show("Fehler: AD Surname ist leer - Verwende Windows Username: " + curentUserName);


                        return curentUserName;
                    } else {
                        return ADsurname;
                    }

                    
                }
            }
            catch (Exception ex)
            {
                lbl_info.Text = DateTime.UtcNow + ": Fehler beim Abfragen des Nachnamens aus der AD. Benutze Windows Username!";
                lbl_info.Text = DateTime.UtcNow + ex.Message;
                return curentUserName;
            }
            // Wenn weder AD User noch lokaler User abgefragt werden kann.
            return null;
        }

        public static string GetProperty(SearchResult searchResult, 
 string PropertyName)
  {
   if(searchResult.Properties.Contains(PropertyName))
   {
    return searchResult.Properties[PropertyName][0].ToString() ;
   }
   else
   {
    return string.Empty;
   }
  }

        private void checkNewTicket_CheckedChanged(object sender, EventArgs e)
        {
            if (checkNewTicket.Checked == false)
            {
                tbox_TicketID.Enabled = true;
                labelTicketID.Enabled = true;
            } else 
            {
                tbox_TicketID.Enabled = false;
                labelTicketID.Enabled = false;
            }
        }




        // ---> Button wieder entfernt, da das Ticket erstellen über Mail sich als umständlich erwiesen hat.
        //              
        //private void btn_newTicket_Click(object sender, EventArgs e)
        //{

        //    DialogResult res = MessageBox.Show("Achtung, Sie wollen ein neues Ticket ertellen! \n\n Bitte in Helpdesk: \n 1: Überprüfen ob der Meldende auch meldeberechtigt ist. \n 2: Die Kundeninformationen des Tickets richtig setzen! \n \n \n Neues Ticket erstellen?", "Achtung!", MessageBoxButtons.YesNo);

        //    if (res == DialogResult.Yes)
        //    {
        //        _param.IsChecked = true;
        //        _param.Subject = txtbox_Subject.Text;
        //        this.Close();
        //    }
        //    else
        //    {
        //        // Nothing
        //    }

        //}





    }
}
