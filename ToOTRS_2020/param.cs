using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestOutlookAddIn
{
    public class Param
    {
        private string _TicketID;

        private string _subject;

        private string _regPath;

        private string _regKey;

        private string _newSubject;

        private string _oldSubject;

        private bool isChecked;

        private string _version;



        /// ///////////////////////////////////////////////

        public string Version
        {
            get { return _version; }
            set { _version = value; }
        }

        public bool IsChecked
        {
            get { return isChecked; }
            set { isChecked = value; }
        }
        public string OldSubject
        {
            get { return _oldSubject; }
            set { _oldSubject = value; }
        }

        public string NewSubject
        {
            get { return _newSubject; }
            set { _newSubject = value; }
        }

        public string RegPath
        {
            get { return _regPath; }
            set { _regPath = value; }
        }

        public string RegKey
        {
            get { return _regKey; }
            set { _regKey = value; }
        }

        public string Subject
        {
            get { return _subject; }
            set { _subject = value; }
        }

        public string TicketID
        {
            get { return _TicketID; }
            set { _TicketID = value; }
        }



    }
}
