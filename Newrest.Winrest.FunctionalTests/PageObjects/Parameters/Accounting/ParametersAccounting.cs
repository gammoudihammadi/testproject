using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Globalization;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Accounting
{
    public class ParametersAccounting : PageBase
    {

        public ParametersAccounting(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        //_________________________________Constantes___________________________________________________

        private const string SAVE = "last";

        // Groups & VATs
        private const string GROUP_VATS_TAB = "groupAndVatsTab";
        private const string ITEM_GROUP_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[(contains(td[1],'{0}') and contains(td[2], '{1}'))]";
        private const string ITEM_GROUP_LINE_SS_VAT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[contains(text(),'{0}')]";
        private const string ITEM_GROUP_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[(contains(td[1],'{0}') and contains(td[2], '{1}'))]/td[7]/a[1]";
        private const string ITEM_GROUP_EDIT_SS_VAT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[contains(text(),'{0}')]/../td[7]/a[1]";
        private const string ITEM_GROUP_DELETE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[(contains(td[1],'{0}') and contains(td[2], '{1}'))]/td[7]/a[2]";
        private const string CONFIRM_DELETE = "first";

        private const string NEW_GROUP = "//*[@id=\"tabContentParameters\"]/div[1]/a";
        private const string ITEM_GROUP_NAME = "//*[@id=\"ItemGroupTaxModal\"]/div/div/div/div/form/div[2]/div[1]/div/select";
        private const string GROUP_VAT_VALUE = "//*[@id=\"ItemGroupTaxModal\"]/div/div/div/div/form/div[2]/div[2]/div/select";
        private const string GROUP_ACCOUNT = "AccountingCode";
        private const string GROUP_EXO_ACCOUNT = "ExoneratedAccountingCode";
        private const string GROUP_INV_ACCOUNT = "InventoryAccount";
        private const string GROUP_INV_VAR_ACCOUNT = "InventoryVariationAccount";

        // Service category & VATs
        private const string SERVICE_CATEGORY_VATS_TAB = "groupCategoriesVatsTab";
        private const string SERVICE_CATEGORY_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[(contains(td[1],'{0}') and contains(td[2], '{1}'))]";
        private const string SERVICE_CATEGORY_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[(contains(td[1],'{0}') and contains(td[2], '{1}'))]/td[6]/a[1]";
        private const string SERVICE_CATEGORY_DELETE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[(contains(td[1],'{0}') and contains(td[2], '{1}'))]/td[6]/a[2]";

        private const string NEW_SERVICE_CATEGORY = "//*[@id=\"tabContentParameters\"]/div[1]/a";
        private const string SERVICE_CATEGORY_NAME = "//*[@id=\"ServiceCategoryTaxModal\"]/div/div/div/div/form/div[2]/div[1]/div/select";
        private const string SERVICE_CATEGORY_VAT_VALUE = "//*[@id=\"ServiceCategoryTaxModal\"]/div/div/div/div/form/div[2]/div[2]/div/select";
        private const string SERVICE_CATEGORY_CUSTOMER_TYPE = "//*[@id=\"ServiceCategoryTaxModal\"]/div/div/div/div/form/div[2]/div[3]/div/select";
        private const string SERVICE_CATEGORY_ACCOUNT = "AccountingCode";
        private const string SERVICE_CATEGORY_EXO_ACCOUNT = "ExoneratedAccountingCode";

        // Airport Fee
        private const string AIRPORT_FEE_TAB = "airportFeeTab";
        private const string ADD_SITE = "//*[@id=\"tabContentParameters\"]/div[1]/a";
        private const string SITE_NAME = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[{0}]/td[1]";
        private const string CUSTOMER_LIST = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[{0}]/td[2]";
        private const string TAX_VALUE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[{0}]/td[3]";
        private const string SITE_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[{0}]/td[4]/a[1]";
        private const string VAR_RAT = "/html/body/div[2]/div/div/div/table/tbody/tr[2]/td[3]";

        // Journal
        private const string JOURNAL_TAB = "journalTab";
        private const string EDIT_JOURNAL_DEV = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[2]/td[10]/a[1]/span";
        private const string EDIT_JOURNAL_PATCH = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[2]/td[8]/a[1]/span";

        // Account settings

        private const string ACCOUNT_SETTINGS_TAB = "accountSettingsTab";
        private const string SAGE_CLOSURE_INDEX_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"SageClosureDayIndex\"]";
        private const string SAGE_CLOSURE_INDEX_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"SageClosureDayIndex\"]/../td[6]/a/span";
        private const string CLOSURE_DAY = "Value";//monthlyClosingDayTab


        private const string MONTHLY_CLOSING_MONTH_TAB = "monthlyClosingDayTab";
        //private const string SAGE_CLOSURE_INDEX_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"SageClosureDayIndex\"]";
        private const string SAGE_CLOSURE_MONTHLY_INDEX_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td";
        private const string CLOSURE_MONTHLY_DAY = "ClosingDayForMonth";//monthlyClosingDayTab

        private const string EXPORT_SAGE_VERSION_LINE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"ExportSageVersion\"]";
        private const string EXPORT_SAGE_VERSION_EDIT = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[3][normalize-space(text())=\"ExportSageVersion\"]/../td[6]/a/span";
        private const string EXPORT_SAGE_VERSION = "SelectedSpecificValue";


        // Redistributive Customer Tax -----------------------------------------------------------------
        private const string REDISTRIBUTIVE_CUSTOMER_TAX_TAB = "redistributiveCustomerTab";
        private const string CREATE_NEW_REDISTRIBUTIVE_CUSTOMER_TAX_ID = "btn_rct_new";

        [FindsBy(How = How.Id, Using = REDISTRIBUTIVE_CUSTOMER_TAX_TAB)]
        private IWebElement _redistributiveCustomerTax_tab;

        [FindsBy(How = How.Id, Using = CREATE_NEW_REDISTRIBUTIVE_CUSTOMER_TAX_ID)]
        private IWebElement _createNewRedistributiveCustomerTax;


        // Redistributive Supplier Tax -----------------------------------------------------------------
        private const string REDISTRIBUTIVE_SUPPLIER_TAX_TAB = "redistributiveSupplierTab";
        private const string CREATE_NEW_REDISTRIBUTIVE_SUPPLIER_TAX_ID = "btn_rst_new";

        [FindsBy(How = How.Id, Using = REDISTRIBUTIVE_SUPPLIER_TAX_TAB)]
        private IWebElement _redistributiveSupplierTax_tab;

        [FindsBy(How = How.Id, Using = CREATE_NEW_REDISTRIBUTIVE_SUPPLIER_TAX_ID)]
        private IWebElement _createNewRedistributiveSupplierTax;

        // ---------------------------------------------------------------------------------------------




        //_________________________________Variables____________________________________________________

        // Group & VATs
        [FindsBy(How = How.Id, Using = GROUP_VATS_TAB)]
        private IWebElement _groupVats_tab;

        [FindsBy(How = How.XPath, Using = NEW_GROUP)]
        private IWebElement _newGroup;

        [FindsBy(How = How.XPath, Using = ITEM_GROUP_NAME)]
        private IWebElement _itemGroup;

        [FindsBy(How = How.XPath, Using = GROUP_VAT_VALUE)]
        private IWebElement _vatValue;

        [FindsBy(How = How.Id, Using = GROUP_ACCOUNT)]
        private IWebElement _account;

        [FindsBy(How = How.Id, Using = GROUP_EXO_ACCOUNT)]
        private IWebElement _exonaratedAccount;

        [FindsBy(How = How.Id, Using = GROUP_INV_ACCOUNT)]
        private IWebElement _inventoryAccount;

        [FindsBy(How = How.Id, Using = GROUP_INV_VAR_ACCOUNT)]
        private IWebElement _inventoryVariationAccount;

        [FindsBy(How = How.Id, Using = SAVE)]
        private IWebElement _save;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        // Service category & VATs
        [FindsBy(How = How.Id, Using = SERVICE_CATEGORY_VATS_TAB)]
        private IWebElement _serviceCategoryVats_tab;

        [FindsBy(How = How.XPath, Using = NEW_SERVICE_CATEGORY)]
        private IWebElement _newServiceCategory;

        [FindsBy(How = How.XPath, Using = SERVICE_CATEGORY_NAME)]
        private IWebElement _serviceCategoryName;

        [FindsBy(How = How.XPath, Using = SERVICE_CATEGORY_VAT_VALUE)]
        private IWebElement _serviceCategoryVatValue;

        [FindsBy(How = How.XPath, Using = SERVICE_CATEGORY_CUSTOMER_TYPE)]
        private IWebElement _serviceCategoryCustomerType;

        [FindsBy(How = How.Id, Using = SERVICE_CATEGORY_ACCOUNT)]
        private IWebElement _serviceCategoryAccount;

        [FindsBy(How = How.Id, Using = SERVICE_CATEGORY_EXO_ACCOUNT)]
        private IWebElement _serviceCategoryExonaratedAccount;

        // AirportFee
        [FindsBy(How = How.Id, Using = AIRPORT_FEE_TAB)]
        private IWebElement _airportFee_tab;

        [FindsBy(How = How.XPath, Using = ADD_SITE)]
        private IWebElement _addSite;

        // Journal     
        [FindsBy(How = How.Id, Using = JOURNAL_TAB)]
        private IWebElement _journal_tab1;
        ///

        [FindsBy(How = How.XPath, Using = EDIT_JOURNAL_DEV)]
        private IWebElement _editJournalDev;

        [FindsBy(How = How.XPath, Using = EDIT_JOURNAL_PATCH)]
        private IWebElement _editJournalPatch;

        // Account settings

        [FindsBy(How = How.Id, Using = MONTHLY_CLOSING_MONTH_TAB)]
        private IWebElement _monthlyClosingMonth_Tab;

        [FindsBy(How = How.XPath, Using = SAGE_CLOSURE_MONTHLY_INDEX_EDIT)]
        private IWebElement _sageClosureMonthlyIndexLine;

        [FindsBy(How = How.Id, Using = CLOSURE_MONTHLY_DAY)]
        private IWebElement _closureMonthlyDay;

        //Obsolete 
        [FindsBy(How = How.Id, Using = ACCOUNT_SETTINGS_TAB)]
        private IWebElement _accountSettings_tab1;

        [FindsBy(How = How.XPath, Using = SAGE_CLOSURE_INDEX_LINE)]
        private IWebElement _sageClosureIndexLine;

        [FindsBy(How = How.XPath, Using = SAGE_CLOSURE_INDEX_EDIT)]
        private IWebElement _sageClosureIndexEdit;

        [FindsBy(How = How.Id, Using = CLOSURE_DAY)]
        private IWebElement _closureDay;
        //

        [FindsBy(How = How.XPath, Using = EXPORT_SAGE_VERSION_LINE)]
        private IWebElement _exportSageVersionLine;

        [FindsBy(How = How.XPath, Using = EXPORT_SAGE_VERSION_EDIT)]
        private IWebElement _exportSageVersionEdit;

        [FindsBy(How = How.Id, Using = EXPORT_SAGE_VERSION)]
        private IWebElement _exportSageVersion;

        //________________________________________ Group & VATs ______________________________________________

        public void GoToTab_GroupVats()
        {
            _groupVats_tab = WaitForElementIsVisible(By.Id(GROUP_VATS_TAB));
            _groupVats_tab.Click();
            WaitForLoad();
        }

        public void SearchGroup(string groupName, string vatValue)
        {
            IWebElement editGroup;

            if (!String.IsNullOrEmpty(vatValue))
            {
                editGroup = WaitForElementExists(By.XPath(string.Format(ITEM_GROUP_EDIT, groupName, vatValue)));
            }
            else
            {
                editGroup = WaitForElementExists(By.XPath(string.Format(ITEM_GROUP_EDIT_SS_VAT, groupName)));
            }

            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(false);", editGroup);
            WaitPageLoading();

            editGroup.Click();
            WaitForLoad();
        }

        public void CreateNewGroup(string groupName, string vatValue)
        {
            _newGroup = WaitForElementIsVisible(By.XPath(NEW_GROUP));
            _newGroup.Click();
            WaitForLoad();

            _itemGroup = WaitForElementIsVisible(By.XPath(ITEM_GROUP_NAME));
            _itemGroup.SetValue(ControlType.DropDownList, groupName);

            _vatValue = WaitForElementIsVisible(By.XPath(GROUP_VAT_VALUE));
            _vatValue.SetValue(ControlType.DropDownList, vatValue);

            _account = WaitForElementIsVisible(By.Id(GROUP_ACCOUNT));
            _account.SetValue(ControlType.TextBox, "0");

            _exonaratedAccount = WaitForElementIsVisible(By.Id(GROUP_EXO_ACCOUNT));
            _exonaratedAccount.SetValue(ControlType.TextBox, "0");

            _inventoryAccount = WaitForElementIsVisible(By.Id(GROUP_INV_ACCOUNT));
            _inventoryAccount.SetValue(ControlType.TextBox, "0");

            _inventoryVariationAccount = WaitForElementIsVisible(By.Id(GROUP_INV_VAR_ACCOUNT));
            _inventoryVariationAccount.SetValue(ControlType.TextBox, "0");

            _save = WaitForElementIsVisible(By.Id(SAVE));
            _save.Click();
            WaitForLoad();
        }

        public void EditInventoryAccounts(string account = null, string exoAccount = null, string invAccount = null, string invVarAccount = null)
        {
            if (account != null)
            {
                _account = WaitForElementIsVisible(By.Id(GROUP_ACCOUNT));
                _account.SetValue(ControlType.TextBox, account);
            }

            if (exoAccount != null)
            {
                _exonaratedAccount = WaitForElementIsVisible(By.Id(GROUP_EXO_ACCOUNT));
                _exonaratedAccount.SetValue(ControlType.TextBox, exoAccount);
            }

            if (invAccount != null)
            {
                _inventoryAccount = WaitForElementIsVisible(By.Id(GROUP_INV_ACCOUNT));
                _inventoryAccount.SetValue(ControlType.TextBox, invAccount);
            }

            if (invVarAccount != null)
            {
                _inventoryVariationAccount = WaitForElementIsVisible(By.Id(GROUP_INV_VAR_ACCOUNT));
                _inventoryVariationAccount.SetValue(ControlType.TextBox, invVarAccount);
            }

            _save = WaitForElementIsVisible(By.Id(SAVE));
            _save.Click();
            WaitForLoad();
        }

        public bool IsGroupAndTaxPresent(string groupName, string taxType)
        {
            try
            {
                _webDriver.FindElement(By.XPath(string.Format(ITEM_GROUP_LINE, groupName, taxType)));
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool IsGroupPresent(string groupName)
        {
            try
            {
                _webDriver.FindElement(By.XPath(string.Format(ITEM_GROUP_LINE_SS_VAT, groupName)));
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void DeleteGroup(string groupName, string vatValue)
        {
            var deleteGroup = WaitForElementExists(By.XPath(string.Format(ITEM_GROUP_DELETE, groupName, vatValue)));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(false);", deleteGroup);
            WaitPageLoading();

            deleteGroup.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitForLoad();
        }

        //________________________________________ Service categories & VATs ______________________________________________

        public void GoToTab_ServiceCategoryVats()
        {
            _serviceCategoryVats_tab = WaitForElementIsVisible(By.Id(SERVICE_CATEGORY_VATS_TAB));
            _serviceCategoryVats_tab.Click();
            WaitForLoad();
        }

        public void SearchServiceCategory(string serviceCategoryName, string vatValue)
        {
            var editCategory = WaitForElementExists(By.XPath(string.Format(SERVICE_CATEGORY_EDIT, serviceCategoryName, vatValue)));

            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(false);", editCategory);
            WaitPageLoading();

            editCategory.Click();
            WaitForLoad();
        }

        public void CreateNewServiceCategory(string serviceCategoryName, string vatValue, string customerType)
        {
            _newServiceCategory = WaitForElementIsVisible(By.XPath(NEW_SERVICE_CATEGORY));
            _newServiceCategory.Click();
            WaitForLoad();

            _serviceCategoryName = WaitForElementIsVisible(By.XPath(SERVICE_CATEGORY_NAME));
            _serviceCategoryName.SetValue(ControlType.DropDownList, serviceCategoryName);

            _serviceCategoryVatValue = WaitForElementIsVisible(By.XPath(SERVICE_CATEGORY_VAT_VALUE));
            _serviceCategoryVatValue.SetValue(ControlType.DropDownList, vatValue);

            _serviceCategoryCustomerType = WaitForElementIsVisible(By.XPath(SERVICE_CATEGORY_CUSTOMER_TYPE));
            _serviceCategoryCustomerType.SetValue(ControlType.DropDownList, customerType);

            _serviceCategoryAccount = WaitForElementIsVisible(By.Id(SERVICE_CATEGORY_ACCOUNT));
            _serviceCategoryAccount.SetValue(ControlType.TextBox, "0");

            _serviceCategoryExonaratedAccount = WaitForElementIsVisible(By.Id(SERVICE_CATEGORY_EXO_ACCOUNT));
            _serviceCategoryExonaratedAccount.SetValue(ControlType.TextBox, "0");

            _save = WaitForElementIsVisible(By.Id(SAVE));
            _save.Click();
            WaitForLoad();
        }

        public void EditInventoryAccounts(string account = null, string exoAccount = null)
        {
            if (account != null)
            {
                _serviceCategoryAccount = WaitForElementIsVisible(By.Id(SERVICE_CATEGORY_ACCOUNT));
                _serviceCategoryAccount.SetValue(ControlType.TextBox, account);
            }

            if (exoAccount != null)
            {
                _serviceCategoryExonaratedAccount = WaitForElementIsVisible(By.Id(SERVICE_CATEGORY_EXO_ACCOUNT));
                _serviceCategoryExonaratedAccount.SetValue(ControlType.TextBox, exoAccount);
            }

            _save = WaitForElementIsVisible(By.Id(SAVE));
            _save.Click();
            WaitForLoad();
        }

        public bool IsServiceCategoryAndTaxPresent(string serviceCategoryName, string taxType)
        {
            try
            {
                _webDriver.FindElement(By.XPath(string.Format(SERVICE_CATEGORY_LINE, serviceCategoryName, taxType)));
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void DeleteCategory(string categoryName, string vatValue)
        {
            var deleteCategory = WaitForElementExists(By.XPath(string.Format(SERVICE_CATEGORY_DELETE, categoryName, vatValue)));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(false);", deleteCategory);
            WaitPageLoading();

            deleteCategory.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitForLoad();
        }

        //________________________________________Journal______________________________________________

        public void GoToTab_Journal()
        {
            _journal_tab1 = WaitForElementIsVisible(By.Id(JOURNAL_TAB));
            _journal_tab1.Click();
            WaitForLoad();
        }

        public void EditJournal(string site, string journalInvoice = null, string journalSI = null, string journalRN = null, string journalInventory = null, string journalOF = null)
        {
            var editJournalPage = OpenJournal();

            if (journalInvoice != null)
            {
                editJournalPage.AddJournalInvoice(journalInvoice);
            }

            editJournalPage.AddJournalDueToInvoiceExcludedVATCode("NT");

            if (journalSI != null)
            {
                editJournalPage.AddJournalPurchase(journalSI);
            }

            if (journalRN != null)
            {
                editJournalPage.AddJournalDueToInvoice(journalRN);
            }

            if (journalInventory != null)
            {
                editJournalPage.AddJournalInventory(journalInventory);
            }

            if (journalOF != null)
            {
                editJournalPage.AddJournalHandOver(journalOF);
            }

            editJournalPage.AddJournalHandOverExculdedVATCode("NT");

            editJournalPage.AddSites(site);
            WaitForLoad();
            WaitPageLoading();
        }

        public void RemoveSiteInJournal(string site)
        {
            var editJournalPage = OpenJournal();
            editJournalPage.RemoveSite(site);
        }

        private ParametersAccountingEditJournalModalPage OpenJournal()
        {
            try
            {
                _editJournalDev = WaitForElementExists(By.XPath(EDIT_JOURNAL_DEV));
                _editJournalDev.Click();
            }
            catch
            {
                _editJournalPatch = WaitForElementExists(By.XPath(EDIT_JOURNAL_PATCH));
                _editJournalPatch.Click();
            }

            WaitForLoad();

            return new ParametersAccountingEditJournalModalPage(_webDriver, _testContext);
        }

        //_________________________________________AIRPORT_FEE___________________________________________

        public void GoToTab_AirportFee()
        {
            _airportFee_tab = WaitForElementIsVisible(By.Id(AIRPORT_FEE_TAB));
            _airportFee_tab.Click();
            WaitForLoad();
        }

        public void SetAirportFeeSiteAndCustomer(string site, string customerCode, string customer, string taxValue)
        {
            bool isSiteFound = false;

            // On regarde s'il y a des sites déjà définis
            var nbLignes = _webDriver.FindElements(By.TagName("tr")).Count;

            if (nbLignes > 1)
            {
                for (int i = 1; i < nbLignes; i++)
                {
                    var xpath = String.Format(SITE_NAME, i + 1);
                    var siteName = WaitForElementIsVisible(By.XPath(xpath));

                    if (siteName.Text.Equals(site))
                    {

                        var isCustomerAllowed = VerifyCustomer(customerCode, i);
                        var isValueDifferentTo0 = VerifyValue(i);

                        if (!isCustomerAllowed || !isValueDifferentTo0)
                        {
                            // On ouvre la page d'édition du site
                            var editSitePage = EditSite(i);

                            if (!isCustomerAllowed)
                            {
                                editSitePage.SetCustomer(customer);
                            }

                            if (!isValueDifferentTo0)
                            {
                                editSitePage.SetValue(taxValue);
                            }

                            editSitePage.Save();
                        }

                        isSiteFound = true;
                        break;
                    }
                }
            }

            if (!isSiteFound)
            {
                isSiteFound = AddSite(site, taxValue);

            }

            Assert.IsTrue(isSiteFound, "La taxe Airport Tax n'a pas pu être assignée au site " + site + ".");
        }

        public bool VerifyCustomer(string customer, int i)
        {
            var customerCode = String.Format(CUSTOMER_LIST, i + 1);
            var customers = WaitForElementIsVisible(By.XPath(customerCode));
            var value = customers.Text;

            return (value.Contains(customer));
        }

        public bool VerifyValue(int i)
        {
            var taxValues = String.Format(TAX_VALUE, i + 1);
            var taxValue = WaitForElementIsVisible(By.XPath(taxValues));
            var value = taxValue.Text;

            return (!value.Equals("0"));
        }

        public ParametersAccountingCreateSiteModalPage EditSite(int i)
        {
            var edit = String.Format(SITE_EDIT, i + 1);
            var editSite = WaitForElementIsVisible(By.XPath(edit));
            editSite.Click();

            return new ParametersAccountingCreateSiteModalPage(_webDriver, _testContext);
        }

        public ParametersAccountingCreateSiteModalPage AddNewSite()
        {
            _addSite = WaitForElementIsVisible(By.XPath(ADD_SITE));
            _addSite.Click();
            WaitForLoad();

            return new ParametersAccountingCreateSiteModalPage(_webDriver, _testContext);
        }

        public bool AddSite(string site, string taxValue)
        {

            bool isSiteAdded = true;

            try
            {
                // Click sur le bouton New
                var createNewSitePage = AddNewSite();

                // Ajout du site, du customer et de la value
                createNewSitePage.SetSite(site);
                createNewSitePage.SetAllCustomer();
                createNewSitePage.SetValue(taxValue);

                var accountingPage = createNewSitePage.AddNew();
            }
            catch
            {
                isSiteAdded = false;
            }

            return isSiteAdded;
        }

        // ___________________________________________ Account Settings _____________________________________________

        public void GoToTab_AccountSettings()
        {

            _accountSettings_tab1 = WaitForElementIsVisible(By.Id(ACCOUNT_SETTINGS_TAB));
            _accountSettings_tab1.Click();
            WaitForLoad();

        }

        public void GoToTab_MonthlyClosingDays()
        {

            _monthlyClosingMonth_Tab = WaitForElementIsVisible(By.Id(MONTHLY_CLOSING_MONTH_TAB));
            _monthlyClosingMonth_Tab.Click();
            WaitForLoad();

        }

        public void SetExportSageVersion(string sageVersion)
        {
            Actions action = new Actions(_webDriver);

            action.MoveToElement(_exportSageVersionLine).Perform();
            _exportSageVersionLine.Click();

            _exportSageVersionEdit = WaitForElementIsVisible(By.XPath(EXPORT_SAGE_VERSION_EDIT));
            _exportSageVersionEdit.Click();
            WaitForLoad();

            _exportSageVersion = WaitForElementIsVisible(By.Id(EXPORT_SAGE_VERSION));
            _exportSageVersion.SetValue(ControlType.DropDownList, sageVersion);

            _save = WaitForElementIsVisible(By.Id(SAVE));
            _save.Click();
            WaitForLoad();
        }

        public DateTime GetSageClosureDayIndex()
        {
            Actions action = new Actions(_webDriver);

            action.MoveToElement(_sageClosureIndexLine).Perform();
            _sageClosureIndexLine.Click();

            _sageClosureIndexEdit = WaitForElementIsVisible(By.XPath(SAGE_CLOSURE_INDEX_EDIT));
            _sageClosureIndexEdit.Click();
            WaitForLoad();

            _closureDay = WaitForElementIsVisible(By.Id(CLOSURE_DAY));

            if (_closureDay.GetAttribute("value").Equals(""))
            {
                _closureDay.SetValue(ControlType.TextBox, 1);
            }

            var integrationDate = new DateTime(DateUtils.Now.Year, DateUtils.Now.Month, Int32.Parse(_closureDay.GetAttribute("value")));

            _save = WaitForElementIsVisible(By.Id(SAVE));
            _save.Click();
            WaitForLoad();

            return integrationDate;
        }


        public DateTime GetSageClosureMonthIndex()
        {
            //Actions action = new Actions(_webDriver);
            var MonthAndYear = DateUtils.Now.ToString("MM/yyyy");

            _sageClosureMonthlyIndexLine = WaitForElementIsVisible(By.XPath("//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]/td[contains(text(),'" + MonthAndYear + "')]/../td[2]/a[1]"));
            _sageClosureMonthlyIndexLine.Click();
            WaitForLoad();

            _closureMonthlyDay = WaitForElementIsVisible(By.Id("AccountingMonthlyClosingDay_ClosingDayForMonth"));

            if (_closureMonthlyDay.GetAttribute("value").Equals(""))
            {
                _closureMonthlyDay.SendKeys(DateUtils.Now.Month.ToString() + "05" + DateUtils.Now.Year.ToString());// "07062021"
            }


            DateTime integrationDate = DateTime.ParseExact(_closureMonthlyDay.GetAttribute("value"), "yyyy-MM-dd", CultureInfo.InvariantCulture);


            _save = WaitForElementIsVisible(By.Id(SAVE));
            _save.Click();
            WaitForLoad();

            return integrationDate;
        }



        //------------ Redistributive Customer/Supplier Tax -------------------------------------------
        public void GoToTab_RedistributiveCustomerTax()
        {
            _redistributiveCustomerTax_tab = WaitForElementIsVisible(By.Id(REDISTRIBUTIVE_CUSTOMER_TAX_TAB));
            _redistributiveCustomerTax_tab.Click();
            WaitForLoad();
        }

        public void GoToTab_RedistributiveSupplierTax()
        {
            _redistributiveSupplierTax_tab = WaitForElementIsVisible(By.Id(REDISTRIBUTIVE_SUPPLIER_TAX_TAB));
            _redistributiveSupplierTax_tab.Click();
            WaitForLoad();
        }



        public ParametersAccountingCreateRedistributiveCustomerTaxModalPage ClickOnNewRedistributiveCustomerTax()
        {
            _createNewRedistributiveCustomerTax = WaitForElementIsVisible(By.Id(CREATE_NEW_REDISTRIBUTIVE_CUSTOMER_TAX_ID));
            _createNewRedistributiveCustomerTax.Click();
            WaitForLoad();

            return new ParametersAccountingCreateRedistributiveCustomerTaxModalPage(_webDriver, _testContext);
        }

        public ParametersAccountingCreateRedistributiveSupplierTaxModalPage ClickOnNewRedistributiveSupplierTax()
        {
            _createNewRedistributiveSupplierTax = WaitForElementIsVisible(By.Id(CREATE_NEW_REDISTRIBUTIVE_SUPPLIER_TAX_ID));
            _createNewRedistributiveSupplierTax.Click();
            WaitForLoad();

            return new ParametersAccountingCreateRedistributiveSupplierTaxModalPage(_webDriver, _testContext);
        }

        public int RedistributiveCustomerTaxCount(string taxtype, string site, string customer, string category, double value)
        {
            int row = 2;
            int count = 0;

            while (true)
            {
                // Échec de Assert.IsTrue. There's already a same site - service - tax customer relation in database.
                ////*[@id="tabContentParameters"]/table/tbody/tr[2]/td[1]
                var row_exists = this.isElementExists(By.XPath($"//*[@id=\"tabContentParameters\"]/table/tbody/tr[{row}]/td[1]"));
                if (!row_exists)
                {
                    break;
                }

                var col1_vattax = this.WaitForElementExists(By.XPath($"//*[@id=\"tabContentParameters\"]/table/tbody/tr[{row}]/td[1]"));
                var vattax_value = col1_vattax?.Text;

                var col2_site = this.WaitForElementExists(By.XPath($"//*[@id=\"tabContentParameters\"]/table/tbody/tr[{row}]/td[2]"));
                var site_value = col2_site?.Text;

                var col3_customers = this.WaitForElementExists(By.XPath($"//*[@id=\"tabContentParameters\"]/table/tbody/tr[{row}]/td[3]"));
                var customers_value = col3_customers?.Text;

                var col4_categories = this.WaitForElementExists(By.XPath($"//*[@id=\"tabContentParameters\"]/table/tbody/tr[{row}]/td[4]"));
                var categories_value = col4_categories?.Text;

                var col5_value = this.WaitForElementExists(By.XPath($"//*[@id=\"tabContentParameters\"]/table/tbody/tr[{row}]/td[5]"));
                var value_value = col5_value?.Text;


                if (vattax_value.Equals(taxtype)
                    && site_value.Equals(site)
                    && customers_value.Equals(customer)
                    && categories_value.Equals(category)
                    && value_value.Equals(value.ToString()))
                {
                    count++;
                    Console.WriteLine($"redistributive customer tax found at index {row}");

                }

                row++;
            }

            return count;
        }

        public int RedistributiveSupplierTaxCount(string taxtype, string site, string supplier, string itemgroup, double value)
        {
            int row = 2;
            int count = 0;

            while (true)
            {
                var row_exists = this.isElementExists(By.XPath($"//*[@id=\"tabContentParameters\"]/table/tbody/tr[{row}]/td[1]"));
                if (!row_exists)
                {
                    break;
                }

                var col1_vattax = this.WaitForElementExists(By.XPath($"//*[@id=\"tabContentParameters\"]/table/tbody/tr[{row}]/td[1]"));
                var vattax_value = col1_vattax?.Text;

                var col2_site = this.WaitForElementExists(By.XPath($"//*[@id=\"tabContentParameters\"]/table/tbody/tr[{row}]/td[2]"));
                var site_value = col2_site?.Text;

                var col3_suppliers = this.WaitForElementExists(By.XPath($"//*[@id=\"tabContentParameters\"]/table/tbody/tr[{row}]/td[3]"));
                var suppliers_value = col3_suppliers?.Text;

                var col4_itemgroups = this.WaitForElementExists(By.XPath($"//*[@id=\"tabContentParameters\"]/table/tbody/tr[{row}]/td[4]"));
                var itemgroups_value = col4_itemgroups?.Text;

                var col5_value = this.WaitForElementExists(By.XPath($"//*[@id=\"tabContentParameters\"]/table/tbody/tr[{row}]/td[5]"));
                var value_value = col5_value?.Text;


                if (vattax_value.Equals(taxtype)
                    && site_value.Equals(site)
                    && suppliers_value.Equals(supplier)
                    && itemgroups_value.Equals(itemgroup)
                    && value_value.Equals(value.ToString()))
                {
                    count++;
                    Console.WriteLine($"redistributive supplier tax found at index {row}");
                }

                row++;
            }

            return count;
        }


        public string GetVATRATFromAirportFee()
        {
            var taxeFree = WaitForElementIsVisible(By.XPath(VAR_RAT));
            return taxeFree.Text;
        }
    }
}

