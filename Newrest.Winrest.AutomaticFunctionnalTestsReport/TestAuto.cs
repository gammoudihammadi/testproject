using System;
using System.Collections.Generic;
using System.Text;

namespace Newrest.Winrest.AutomaticFunctionnalTestsReport
{
    public class TestAuto
    {
        private string _nomTest;
        private string _environnement;
        private string _status;
        private string _errorMessage;


        public string NomTest
        {
            get { return _nomTest; }
            set { _nomTest = value; }
        }

        public string Environnement
        {
            get { return _environnement; }
            set { _environnement = value; }
        }

        public string Status
        {
            get { return _status; }
            set { _status = value; }
        }

        public string ErrorMessage
        {
            get { return _errorMessage; }
            set { _errorMessage = value; }
        }

    }
}
