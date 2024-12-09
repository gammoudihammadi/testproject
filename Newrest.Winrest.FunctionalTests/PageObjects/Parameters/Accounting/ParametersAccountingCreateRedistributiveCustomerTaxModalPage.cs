using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Accounting
{
    public class ParametersAccountingCreateRedistributiveCustomerTaxModalPage : ParametersAccountingCreateRedistributiveTaxModalPageBase
    {
        public ParametersAccountingCreateRedistributiveCustomerTaxModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        /*
         * 
         * 
delete from RedistributiveTaxCustomers where RedistributiveTax_Id = 1
delete from RedistributiveTaxServiceCategories where RedistributiveTax_Id = 1
delete from RedistributiveTaxes where id = 1
         * 
         * 
         * 
         */

        private const string SELECTED_CUSTOMERS_ID = "SelectedCustomers_ms";
        private const string SELECTED_SERVICE_CATEGORIES_ID = "SelectedServiceCategories_ms";

        private const string VALIDATION_ERROR_MESSAGE_XPATH = "//*[@id=\"RedistributiveCustomerTaxModal\"]/div/div/div/div/div/div/form/div[2]/div[2]/div/div/span";

        public void CreateNewRedistributiveCustomerTax(string taxtype, string site, string customer, string service, double value, string codeJournal = "", string accountCode = "")
        {
            this.DropdownListSelectById(TAXTYPE_ID, taxtype);

            this.ComboBoxSelectById(new ComboBoxOptions(SELECTED_SITES_ID, site, false));
            this.ComboBoxSelectById(new ComboBoxOptions(SELECTED_CUSTOMERS_ID, customer, false));
            this.ComboBoxSelectById(new ComboBoxOptions(SELECTED_SERVICE_CATEGORIES_ID, service, false));


            // fill mandatory fields (name + code)
            var tbValue = this.WaitForElementIsVisible(By.Id(VALUE_ID));
            tbValue.SetValue(PageBase.ControlType.TextBox, value.ToString());

            var tbCodeJournal = this.WaitForElementIsVisible(By.Id(CODE_JOURNAL_ID));
            tbCodeJournal.SetValue(PageBase.ControlType.TextBox, codeJournal);

            var tbAccountCode = this.WaitForElementIsVisible(By.Id(ACCOUNT_CODE_ID));
            tbAccountCode.SetValue(PageBase.ControlType.TextBox, accountCode);

            var btAdd = this.WaitForElementIsVisible(By.Id(ADD_ID));
            btAdd.Click();

            WaitPageLoading();
            WaitForLoad();

            var hasValidationError = this.isElementExists(By.XPath(VALIDATION_ERROR_MESSAGE_XPATH));
            if (hasValidationError)
            {
                var validationError = this.WaitForElementIsVisible(By.XPath(VALIDATION_ERROR_MESSAGE_XPATH));
                Console.WriteLine(validationError.Text);
                throw new Exception(validationError.Text);
            }
        }
    }
}
