using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Accounting
{
    public class ParametersAccountingEditJournalModalPage : PageBase
    {

        public ParametersAccountingEditJournalModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        // ___________________________________________ Constantes ___________________________________________

        private const string JOURNAL_INVOICE = "AccountingJournalSetting_JournalInvoice";
        private const string JOURNAL_INVOICE_EXCLUDED_VAT_CODE = "AccountingJournalSetting_DueToInvoiceExcludedCode";
        private const string JOURNAL_PURCHASE = "AccountingJournalSetting_JournalPurchase";
        private const string JOURNAL_DUE_TO_INVOICE = "AccountingJournalSetting_JournalDueToInvoice";
        private const string JOURNAL_INVENTORY = "AccountingJournalSetting_JournalInventory";
        private const string JOURNAL_HANDOVER = "AccountingJournalSetting_JournalHandOver";
        private const string JOURNAL_HANDOVER_EXCLUDED_VAT_CODE = "AccountingJournalSetting_HandOverExcludedCode";
        private const string SITE_FILTER = "SelectedSites_ms";
        private const string SELECT_ALL = "/html/body/div[11]/div/ul/li[1]/a";
        private const string SEARCH_SITE = "/html/body/div[11]/div/div/label/input";
        private const string SAVE = "last";

        // ___________________________________________ Variables ____________________________________________

        [FindsBy(How = How.Id, Using = JOURNAL_INVOICE)]
        private IWebElement _journalInvoice;

        [FindsBy(How = How.Id, Using = JOURNAL_INVOICE_EXCLUDED_VAT_CODE)]
        private IWebElement _journalInvoiceExcludedVATCode;

        [FindsBy(How = How.Id, Using = JOURNAL_PURCHASE)]
        private IWebElement _journalPurchase;

        [FindsBy(How = How.Id, Using = JOURNAL_DUE_TO_INVOICE)]
        private IWebElement _journalDueToInvoice;

        [FindsBy(How = How.Id, Using = JOURNAL_INVENTORY)]
        private IWebElement _journalInventory;

        [FindsBy(How = How.Id, Using = JOURNAL_HANDOVER)]
        private IWebElement _journalHandOver;

        [FindsBy(How = How.Id, Using = JOURNAL_HANDOVER_EXCLUDED_VAT_CODE)]
        private IWebElement _journalHandOverExcludedVATCode;

        [FindsBy(How = How.Id, Using = SITE_FILTER)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = SELECT_ALL)]
        private IWebElement _selectAll;

        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSite;

        [FindsBy(How = How.Id, Using = SAVE)]
        private IWebElement _save;

        // ___________________________________________ Méthodes _____________________________________________

        public bool AddSites(string site)
        {
            bool isSiteOK = true;
            ComboBoxSelectById(new ComboBoxOptions(SITE_FILTER, site, false));
            
            var nbSite = _webDriver.FindElements((By.XPath("//span[text()='" + site + "']"))).Count;

            if(nbSite == 0)
            {
                isSiteOK = false;
            }

            _save = WaitForElementIsVisible(By.Id(SAVE));
            _save.Click();

            WaitForLoad();

            return isSiteOK;
        }

        public void RemoveSite(string site)
        {
            _siteFilter = WaitForElementIsVisible(By.Id(SITE_FILTER));
            _siteFilter.Click();

            // On sélectionne tous les sites
            _selectAll = WaitForElementIsVisible(By.XPath(SELECT_ALL));
            _selectAll.Click();

            _searchSite = WaitForElementIsVisible(By.XPath(SEARCH_SITE));
            _searchSite.SetValue(ControlType.TextBox, site);
            WaitForLoad();

            var siteFilterValue = WaitForElementIsVisible(By.XPath("//span[text()='" + site + "']"));
            siteFilterValue.Click();

            _save = WaitForElementIsVisible(By.Id(SAVE));
            _save.Click();

            WaitForLoad();
        }

        public void AddJournalInvoice(string journalInvoice)
        {
            _journalInvoice = WaitForElementIsVisible(By.Id(JOURNAL_INVOICE));
            _journalInvoice.SetValue(ControlType.TextBox, journalInvoice);
        }
        public void AddJournalDueToInvoiceExcludedVATCode(string journalInvoiceExcludedVATCode)
        {
            try
            {
                _journalInvoiceExcludedVATCode = WaitForElementIsVisible(By.Id(JOURNAL_INVOICE_EXCLUDED_VAT_CODE));
                _journalInvoiceExcludedVATCode.SetValue(ControlType.TextBox, journalInvoiceExcludedVATCode);
            }
            catch
            {
                // parameter doens't exist on 2021.1109.1-P9
            }
        }

        public void AddJournalPurchase(string journalPurchase)
        {
            _journalPurchase = WaitForElementIsVisible(By.Id(JOURNAL_PURCHASE));
            _journalPurchase.SetValue(ControlType.TextBox, journalPurchase);
        }

        public void AddJournalDueToInvoice(string journalRN)
        {
            _journalDueToInvoice = WaitForElementIsVisible(By.Id(JOURNAL_DUE_TO_INVOICE));
            _journalDueToInvoice.SetValue(ControlType.TextBox, journalRN);
        }

        public void AddJournalInventory(string journalInventory)
        {
            _journalInventory = WaitForElementIsVisible(By.Id(JOURNAL_INVENTORY));
            _journalInventory.SetValue(ControlType.TextBox, journalInventory);
        }

        public void AddJournalHandOver(string journalOF)
        {
            _journalHandOver = WaitForElementIsVisible(By.Id(JOURNAL_HANDOVER));
            _journalHandOver.SetValue(ControlType.TextBox, journalOF);
        }
        public void AddJournalHandOverExculdedVATCode(string journalHandOverExculdedVATCode)
        {
            try
            {
                _journalHandOverExcludedVATCode = WaitForElementIsVisible(By.Id(JOURNAL_HANDOVER_EXCLUDED_VAT_CODE));
                _journalHandOverExcludedVATCode.SetValue(ControlType.TextBox, journalHandOverExculdedVATCode);
            }
            catch
            {
                // parameter doens't exist on 2021.1109.1-P9
            }
        }
    }
}
