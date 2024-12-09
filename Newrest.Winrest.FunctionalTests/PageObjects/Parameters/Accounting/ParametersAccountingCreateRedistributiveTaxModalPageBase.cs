using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Accounting
{
    public abstract class ParametersAccountingCreateRedistributiveTaxModalPageBase : PageBase
    {
        protected ParametersAccountingCreateRedistributiveTaxModalPageBase(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        protected const string TAXTYPE_ID = "TaxTypeId";
        protected const string SELECTED_SITES_ID = "SelectedSites_ms";

        protected const string BY_PERCENT_XPATH = "//*[@id=\"False\"]";
        protected const string BY_VALUE_XPATH = "//*[@id=\"True\"]";

        protected const string VALUE_ID = "Value";
        protected const string CODE_JOURNAL_ID = "CodeJournal";
        protected const string ACCOUNT_CODE_ID = "AccountCode";

        protected const string ADD_ID = "last";

     }
}
