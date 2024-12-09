using FluentAssertions.Execution;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Numeric;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites
{
    public class ParametersSites : PageBase
    {

        public ParametersSites(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        // __________________________________________ Constantes _________________________________________

        // Sites

        // Sites --> General
        private const string SITE_TAB = "href-tab-IndexSites";
        private const string CREATE_NEW_SITE = "//*[@id=\"item-filter-form\"]/div[1]/div[2]/div/a";
        private const string CREATE_NEW_SITE_DEV = "//*[@id=\"site-search\"]/div/div[2]/div/a";
        private const string PLUS_BTN = "//*[@id=\"item-filter-form\"]/div[1]/div[2]/button";
        private const string PLUS_BTN_DEV = "//*[@id=\"site-search\"]/div/div[2]/button";

        // Sites --> Liste sites
        private const string FILTER_SEARCH = "tbSearchPattern";
        private const string PARAMETERS_SITE_EXISTS = "//*[@id=\"tabContentSites\"]/div/div[1]/div[2]/table/tbody/tr[*]/td[2][contains(text(),'{0}')]";
        private const string PARAMETERS_SITE_EXISTS_DEV = "//*[@id=\"item-filter-form\"]/div/div[2]/table/tbody/tr[2]/td[2][contains(text(),'{0}')]";
        private const string ID_RESULT = "//*[@id=\"tabContentSites\"]/div/div[1]/div[2]/table/tbody/tr[2]/td[1]";
        private const string ID_RESULT_DEV = "//*[@id=\"item-filter-form\"]/div/div[2]/table/tbody/tr[2]/td[1]";
        private const string FIRST_SITE = "//*[@id='item-filter-form']/div/div[2]/table/tbody/tr[2]";
        private const string FIRST_SITE_PATCH = "//*[@id=\"tabContentSites\"]/div/div[1]/div[2]/table/tbody/tr[2]";
        private const string IS_ACTIVE = "item_FullSite_IsActive";
        private const string SELECT_ALL = "select-all";
        private const string DEACTIVATE_SITE = "//*[@id=\"deactivate-site-form\"]/div[3]/button[3]";
        private const string FIRST_SITE_NAME = "//*[@id=\"item-filter-form\"]/div/div[2]/table/tbody/tr[2]/td[2]";

        // Sites --> Informations
        private const string INFORMATIONS_SUB_TAB = "//*[@id=\"site-details-panel\"]/div/ul/li[*]/a[text()='Informations']";
        private const string SITE_ADDRESS = "FullSite_Address";
        private const string ANALYTIC_PLAN = "FullSite_AccountingAnalyticPlan";
        private const string ANALYTIC_SECTION = "FullSite_AccountingAnalyticSection";
        private const string DUE_INVOICE_ANALYTIC_PLAN = "FullSite_DueToInvoiceAccountingAnalyticPlan";
        private const string DUE_INVOICE_ANALYTIC_SECTION = "FullSite_DueToInvoiceAccountingAnalyticSection";
        private const string SAGE_AUTO_ENABLED = "FullSite_EnableSageAutoForSite";
        private const string SITE_SANITARY_AGREEMENT = "FullSite_AgreementSanitaryNumber";
        private const string COUNTRY = "//*[@id=\"FullSite_CountryId\"]/option[@selected]";
        private const string SITE_ADDRESS2 = "FullSite_Address2";
        private const string SITE_ADDRESS3 = "FullSite_Address3";
        private const string ZIP_CODE = "FullSite_ZipCode";
        private const string CITY = "FullSite_City";
        private const string SAFETY_NOTE = "input-text-html-body-code-safety";
        private const string SANITARY_DATE = "sanitary-date-picker";
        private const string AIRPORT_NUMBER = "//*[@id=\"FullSite_AgreementAirportNumber\"]";
        private const string EDI_GLN_CODE = "FullSite_CodeSiteEdiGLN";



        // Sites --> Contacts
        private const string CONTACTS_SUB_TAB = "//*[@id=\"site-details-panel\"]/div/ul/li[*]/a[text()='Contacts']";
        private const string INVOICE_SAGE_MANAGER_NAME = "ContactInvoiceSageManager_Name";
        private const string INVOICE_SAGE_MANAGER_MAIL = "ContactInvoiceSageManager_Mail";
        private const string SUPPLIER_INVOICE_SAGE_MANAGER_NAME = "ContactSupplierInvoiceSageManager_Name";
        private const string SUPPLIER_INVOICE_SAGE_MANAGER_MAIL = "ContactSupplierInvoiceSageManager_Mail";
        private const string COMMERCIAL_MANAGER_MAIL = "ContactCommercialManager_Mail";
        private const string COMMERCIAL_MANAGER_NAME = "ContactCommercialManager_Name";

        // Sites --> Airports
        private const string AIRPORT_TAB = "//*[@id=\"site-details-panel\"]/div/ul/li[5]/a";
        private const string AIRPORT = "//*[@id=\"tabContentSiteDetails\"]/table/tbody/tr[*]/td[1]";//*[@id="tabContentSiteDetails"]/table/tbody/tr[1]/td[1]

        // Sites --> Organization
        private const string ORGANIZATION_SUB_TAB = "//*[@id=\"site-details-panel\"]/div/ul/li[3]/a[text()='Organization']";
        private const string NEW_ORGANIZATION = "//*[@id=\"tabContentSiteDetails\"]/div[1]/div[3]/a";
        private const string ORGANIZATION_NAME = "SitePlace_Title";
        private const string CONTACT_NAME = "SitePlace_Contact_Name";
        private const string ORGANIZATION_ACTIVE = "SitePlace_IsActive";
        private const string ORGA_PROD_PLACE = "//*[@id=\"Type\"][@value=\"Production\"]";
        private const string ORGA_STOR_PLACE = "//*[@id=\"Type\"][@value=\"Storage\"]";
        private const string ORGA_SITE_PLACE = "//*/input[@value='Site']";
        private const string SELECT_SITE_PLACE = "sites-list";
        private const string SAVE_ORGANIZATION = "//*[@id=\"modal-1\"]/div/div/div[2]/div/form/div[2]/button[2]";
        private const string SAVE_ORGANIZATION_DEV = "//*[@id=\"modal-1\"]/div[2]/div/form/div[2]/button[2]";

        private const string ORGANIZATION_LINE = "//*[@id=\"tabContentSiteDetails\"]/div[2]/div/div[2]/ul/li[*]/div/div/div[1]/span[text()= '{0}']";
        private const string ORGANIZATION_LINE_DEV = "//*[@id=\"tabContentSiteDetails\"]/div[2]/div/div[2]/ul/li[*]/div/div/div[1]/div/span[text()= '{0}']";
        private const string ORGANIZATION_EDIT = "//*[@id=\"tabContentSiteDetails\"]/div[2]/div/div[2]/ul/li[*]/div/div/div[1]/span[text()= '{0}']/../../div[9]/a";
        private const string ORGANIZATION_EDIT_DEV = "//*[@id=\"tabContentSiteDetails\"]/div[2]/div/div[2]/ul/li[*]/div/div[1]/div/div/span[text()= '{0}']/../../div[9]/a";

        private const string COUNTERS_TAB = "//*[@id=\"site-details-panel\"]/div/ul/li[*]/a[text()='Counters']";
        private const string COUNTERS = "//*[@id=\"table-counters\"]/tbody/tr[*]/td[7]";
        private const string CURRENCY = "FullSite_CurrencyId";
        // Fiscal entities
        private const string FISCAL_ENTITIES_TAB = "href-tab-IndexFiscalEntities";
        private const string FISCAL_ENTITIES_GROUP_EDIT = "//*[@id=\"tabContentSites\"]/table/tbody/tr[*]/td[2][normalize-space(text())='{0}']/../td[4]/a[1]/span";
        private const string FISCAL_NETITY_SAGE_CONNECTION_STRING = "SageConnectionString";
        private const string FISCAL_ENTITY_PLUS_NEW = "//*[@id=\"tabContentSites\"]/a";
        private const string PURCHASE_ORDER_MAILBODY_PLUS_NEW = "//*[@id=\"tabContentSites\"]/div[1]/a";
        private const string PURCHASE_ORDER_MAIL_BODY_TAB = "//*[@id=\"href-tab-PurchaseOrderMailBody\"]/a";
        private const string DELETE_PURCHASE_ORDER_MAILBODY = "//*[@id=\"tabContentSites\"]/table/tbody/tr[2]/td[4]/a[2]";
        private const string DELETE_PURCHASE_BUTTON = "dataConfirmOK";
        private const string INVOICE_MANAGER_MAIL = "/html/body/div[2]/div/div/div/div/div[2]/div/div/div/div/div/form/div[8]/div[3]/div/div/input";
        private const string INVOICE_MANAGER_NAME = "/html/body/div[2]/div/div/div/div/div[2]/div/div/div/div/div/form/div[8]/div[1]/div/div/input";

        private const string CLAIM_MANAGER_MAIL = "ContactClaimManager_Mail";
        private const string CUSTOMER = "CustomersTabNav";
        private const string ORDERTYPE = "//*[@id=\"paramCustomerTab\"]/li[3]/a";
        private const string VERIFORDERTYPE = "//*[@id=\"tabContentParameters\"]/table/tbody/tr[*]";






        // __________________________________________ Variables __________________________________________

        // Sites

        // Sites --> Général
        [FindsBy(How = How.XPath, Using = VERIFORDERTYPE)]
        private IWebElement _veriordertype;
        [FindsBy(How = How.Id, Using = CITY)]
        private IWebElement _city;
        [FindsBy(How = How.XPath, Using = ORDERTYPE)]
        private IWebElement _ordertype;
        [FindsBy(How = How.Id, Using = CUSTOMER)]
        private IWebElement _customer;

        [FindsBy(How = How.Id, Using = ZIP_CODE)]
        private IWebElement _zipcode;

        [FindsBy(How = How.Id, Using = SITE_ADDRESS3)]
        private IWebElement _adress3;

        [FindsBy(How = How.Id, Using = SITE_ADDRESS2)]
        private IWebElement _adress2;

        [FindsBy(How = How.Id, Using = SITE_TAB)]
        private IWebElement _sitesTab;

        [FindsBy(How = How.XPath, Using = PLUS_BTN)]
        private IWebElement _showPlusMenuBtn;

        [FindsBy(How = How.XPath, Using = CREATE_NEW_SITE)]
        private IWebElement _createNewSite;

        // Sites --> Liste sites
        [FindsBy(How = How.Id, Using = FILTER_SEARCH)]
        private IWebElement _searchSite;

        [FindsBy(How = How.XPath, Using = ID_RESULT)]
        private IWebElement _idResult;

        [FindsBy(How = How.Id, Using = IS_ACTIVE)]
        private IWebElement _isActive;

        [FindsBy(How = How.Id, Using = SELECT_ALL)]
        private IWebElement _selectAll;

        [FindsBy(How = How.XPath, Using = DEACTIVATE_SITE)]
        private IWebElement _deactivateSite;

        [FindsBy(How = How.XPath, Using = FIRST_SITE)]
        private IWebElement _firstSite;

        // Sites --> Informations
        [FindsBy(How = How.XPath, Using = INFORMATIONS_SUB_TAB)]
        private IWebElement _informationsTab;

        [FindsBy(How = How.Id, Using = SITE_ADDRESS)]
        private IWebElement _siteAddress;

        [FindsBy(How = How.Id, Using = ANALYTIC_PLAN)]
        private IWebElement _analyticPlan;

        [FindsBy(How = How.Id, Using = ANALYTIC_SECTION)]
        private IWebElement _analyticSection;

        [FindsBy(How = How.Id, Using = DUE_INVOICE_ANALYTIC_PLAN)]
        private IWebElement _dueToInvoiceAnalyticPlan;

        [FindsBy(How = How.Id, Using = DUE_INVOICE_ANALYTIC_SECTION)]
        private IWebElement _dueToInvoiceAnalyticSection;

        [FindsBy(How = How.Id, Using = SAGE_AUTO_ENABLED)]
        private IWebElement _sageAutoEnabled;

        [FindsBy(How = How.Id, Using = SITE_SANITARY_AGREEMENT)]
        private IWebElement _siteSanitaryAgreement;

        [FindsBy(How = How.Id, Using = EDI_GLN_CODE)]
        private IWebElement _ediGNLCode;

        // Sites --> Airport
        [FindsBy(How = How.XPath, Using = AIRPORT_TAB)]
        private IWebElement _aiportTab;

        // Sites --> Contacts
        [FindsBy(How = How.XPath, Using = CONTACTS_SUB_TAB)]
        private IWebElement _contactsTab;

        [FindsBy(How = How.Id, Using = INVOICE_SAGE_MANAGER_NAME)]
        private IWebElement _invoiceSageManagerName;

        [FindsBy(How = How.Id, Using = INVOICE_SAGE_MANAGER_MAIL)]
        private IWebElement _invoiceSageManagerMail;
        [FindsBy(How = How.XPath, Using = INVOICE_MANAGER_MAIL)]
        private IWebElement _invoiceManagerMail;
        [FindsBy(How = How.XPath, Using = INVOICE_MANAGER_NAME)]
        private IWebElement _invoiceManagerName;

        [FindsBy(How = How.Id, Using = SUPPLIER_INVOICE_SAGE_MANAGER_NAME)]
        private IWebElement _supplierInvoiceSageManagerName;

        [FindsBy(How = How.Id, Using = SUPPLIER_INVOICE_SAGE_MANAGER_MAIL)]
        private IWebElement _supplierInvoiceSageManagerMail;

        [FindsBy(How = How.Id, Using = COMMERCIAL_MANAGER_MAIL)]
        private IWebElement _commercialManagerMail;

        [FindsBy(How = How.Id, Using = COMMERCIAL_MANAGER_NAME)]
        private IWebElement _commercialManagerName;

        // Sites --> Organisation
        [FindsBy(How = How.XPath, Using = ORGANIZATION_SUB_TAB)]
        private IWebElement _organizationTab;

        [FindsBy(How = How.XPath, Using = NEW_ORGANIZATION)]
        private IWebElement _newOrganization;

        [FindsBy(How = How.Id, Using = ORGANIZATION_NAME)]
        private IWebElement _organizationName;

        [FindsBy(How = How.Id, Using = CONTACT_NAME)]
        private IWebElement _contactName;

        [FindsBy(How = How.Id, Using = ORGANIZATION_ACTIVE)]
        private IWebElement _isActiveOrganization;

        [FindsBy(How = How.Id, Using = ORGA_PROD_PLACE)]
        private IWebElement _productionPlace;

        [FindsBy(How = How.Id, Using = ORGA_STOR_PLACE)]
        private IWebElement _storagePlace;

        [FindsBy(How = How.XPath, Using = ORGA_SITE_PLACE)]
        private IWebElement _otherSitePlace;

        [FindsBy(How = How.Id, Using = SELECT_SITE_PLACE)]
        private IWebElement _selectSitePlace;

        [FindsBy(How = How.XPath, Using = SAVE_ORGANIZATION)]
        private IWebElement _saveOrganization;


        // Fiscal entities
        [FindsBy(How = How.Id, Using = FISCAL_ENTITIES_TAB)]
        private IWebElement _fiscalEntitiesTab;

        [FindsBy(How = How.XPath, Using = FISCAL_ENTITIES_GROUP_EDIT)]
        private IWebElement _fiscalEntitiesGroupEdit;

        [FindsBy(How = How.Id, Using = FISCAL_NETITY_SAGE_CONNECTION_STRING)]
        private IWebElement _fiscalEntitySageConnectionString;

        [FindsBy(How = How.XPath, Using = FISCAL_ENTITY_PLUS_NEW)]
        private IWebElement _createNewFiscalEntity;

        [FindsBy(How = How.XPath, Using = COUNTERS_TAB)]
        private IWebElement _counterstab;
        [FindsBy(How = How.XPath, Using = PURCHASE_ORDER_MAILBODY_PLUS_NEW)]
        private IWebElement _purchase_order_mailbody;
        [FindsBy(How = How.XPath, Using = PURCHASE_ORDER_MAIL_BODY_TAB)]
        private IWebElement _purchase_order_mail_body;
        [FindsBy(How = How.XPath, Using = DELETE_PURCHASE_ORDER_MAILBODY)]
        private IWebElement _delete_purchase_order_mail_body;
        [FindsBy(How = How.XPath, Using = DELETE_PURCHASE_BUTTON)]
        private IWebElement _delete_purchase_button;


        [FindsBy(How = How.Id, Using = CLAIM_MANAGER_MAIL)]
        private IWebElement _claimManagerMail;

        // _________________________________________ Méthodes ______________________________________________

        // Sites 

        // Sites --> Général

        public void ClickOnSitesTab()
        {
            _sitesTab = WaitForElementIsVisible(By.Id(SITE_TAB));
            _sitesTab.Click();
            WaitForLoad();
        }
        public void ClickOnCustomer()
        {
            _customer = WaitForElementIsVisible(By.Id(CUSTOMER));
            _customer.Click();
            WaitForLoad();
        }
        public void ClickOnOrderType()
        {
            _ordertype = WaitForElementIsVisible(By.XPath(ORDERTYPE));
            _ordertype.Click();
            WaitForLoad();
        }
        public enum FilterType
        {
            SearchSite
        }

        public void Filter(FilterType filterType, object value)
        {
            switch (filterType)
            {
                case FilterType.SearchSite:
                    _searchSite = WaitForElementIsVisible(By.Id(FILTER_SEARCH));
                    _searchSite.SetValue(ControlType.TextBox, value);
                    WaitLoading();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }

            WaitPageLoading();
            WaitForLoad();
        }
        public bool isSiteExists(string site)
        {
            if (isElementVisible(By.XPath(string.Format(PARAMETERS_SITE_EXISTS_DEV, site))))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public override void ShowPlusMenu()
        {
            Actions action = new Actions(_webDriver);
            _showPlusMenuBtn = WaitForElementIsVisible(By.XPath(PLUS_BTN_DEV), nameof(PLUS_BTN_DEV));
            action.MoveToElement(_showPlusMenuBtn).Perform();
        }


        public ParametersSitesModalPage ClickOnNewSite()
        {
            ShowPlusMenu();
            _createNewSite = WaitForElementIsVisible(By.XPath(CREATE_NEW_SITE_DEV));
            _createNewSite.Click();
            WaitForLoad();

            return new ParametersSitesModalPage(_webDriver, _testContext);
        }

        public void ClickOnFirstSite()
        {
            _firstSite = WaitForElementIsVisible(By.XPath(FIRST_SITE));
            _firstSite.Click();
           // WaitLoading();
            WaitPageLoading();
        }

        public string CollectNewSiteID()
        {
            _idResult = WaitForElementIsVisible(By.XPath(ID_RESULT_DEV));
            return _idResult.GetAttribute("innerText");
        }

        public void Deactivate()
        {
            _isActive = WaitForElementExists(By.Id(IS_ACTIVE));
            _isActive.Click();
            WaitForLoad();

            _selectAll = WaitForElementIsVisible(By.Id(SELECT_ALL));
            _selectAll.Click();

            _deactivateSite = WaitForElementIsVisible(By.XPath(DEACTIVATE_SITE));
            _deactivateSite.Click();
            WaitForLoad();
        }

        // Sites --> Informations
        public void ClickToInformations()
        {
            _informationsTab = WaitForElementIsVisible(By.XPath(INFORMATIONS_SUB_TAB));
            _informationsTab.Click();
            WaitForLoad();
            WaitPageLoading();
        }

        public void ClickToAirportTab()
        {
            _aiportTab = WaitForElementIsVisible(By.XPath(AIRPORT_TAB));
            _aiportTab.Click();
            WaitForLoad();
        }

        public List<string> GetAllAirport()
        {
            List<string> airports = new List<string>();
            var elements = _webDriver.FindElements(By.XPath(AIRPORT));

            foreach (var elm in elements)
            {
                airports.Add(elm.Text);
            }

            return airports;
        }

        public List<string> GetAllCounterst()
        {
            List<string> couters = new List<string>();
            var elements = _webDriver.FindElements(By.XPath(COUNTERS));

            foreach (var elm in elements)
            {
                couters.Add(elm.GetAttribute("innerText"));
            }

            return couters;
        }

        public string GetAddress()
        {
            _siteAddress = WaitForElementIsVisible(By.Id(SITE_ADDRESS));
            return _siteAddress.GetAttribute("value");
        }
        public bool VerifOrderType(string name)
        {
             var rows = _webDriver.FindElements(By.XPath("//*[@id='tabContentParameters']/table/tbody/tr[*]"));

            if (rows == null || rows.Count == 0)
            {
                return false;  
            }

             for (int i = 1; i < rows.Count; i++)  
            {
                var row = rows[i];

                 var cells = row.FindElements(By.XPath("./td"));

                 foreach (var cell in cells)
                {
                     var cellValue = cell.Text;

                     if (cellValue.Contains(name)) 
                    {
                        return true;  
                    }
                }
            }

            return false;  
        }






        public string GetCity()
        {
            _city = WaitForElementIsVisible(By.Id(CITY));
            return _city.GetAttribute("value");
        }

        public string GetZipCode()
        {
            _zipcode = WaitForElementIsVisible(By.Id(ZIP_CODE));
            return _zipcode.GetAttribute("value");
        }

        public string GetAddress3()
        {
            _adress3 = WaitForElementIsVisible(By.Id(SITE_ADDRESS3));
            return _adress3.GetAttribute("value");
        }

        public string GetAddress2()
        {
            _adress2 = WaitForElementIsVisible(By.Id(SITE_ADDRESS2));
            return _adress2.GetAttribute("value");
        }

        public void SetAddress(string value)
        {
            _siteAddress = WaitForElementIsVisible(By.Id(SITE_ADDRESS));
            _siteAddress.SetValue(ControlType.TextBox, value);
            WaitPageLoading();
        }

        public string GetAnalyticPlan()
        {
            _analyticPlan = WaitForElementIsVisible(By.Id(ANALYTIC_PLAN));
            return _analyticPlan.GetAttribute("value");
        }

        public void SetAnalyticPlan(string value)
        {
            _analyticPlan = WaitForElementIsVisible(By.Id(ANALYTIC_PLAN));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _analyticPlan);
            _analyticPlan.SetValue(ControlType.TextBox, value);
            WaitPageLoading();
        }

        public string GetAnalyticSection()
        {
            _analyticSection = WaitForElementIsVisible(By.Id(ANALYTIC_SECTION));
            return _analyticSection.GetAttribute("value");
        }

        public void SetAnalyticSection(string value)
        {
            _analyticSection = WaitForElementIsVisible(By.Id(ANALYTIC_SECTION));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _analyticSection);
            _analyticSection.SetValue(ControlType.TextBox, value);
            WaitPageLoading();
        }

        public string GetDueToInvoiceAnalyticSection()
        {
            _dueToInvoiceAnalyticSection = WaitForElementExists(By.Id(DUE_INVOICE_ANALYTIC_SECTION));
            return _dueToInvoiceAnalyticSection.GetAttribute("value");
        }

        public void SetDueToInvoiceAnalyticSection(string value)
        {
            _dueToInvoiceAnalyticSection = WaitForElementExists(By.Id(DUE_INVOICE_ANALYTIC_SECTION));

            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _dueToInvoiceAnalyticSection);
            WaitForLoad();
            _dueToInvoiceAnalyticSection.SetValue(ControlType.TextBox, value);
            WaitPageLoading();
        }

        public string GetDueToInvoiceAnalyticPlan()
        {
            _dueToInvoiceAnalyticPlan = WaitForElementExists(By.Id(DUE_INVOICE_ANALYTIC_PLAN));
            return _dueToInvoiceAnalyticPlan.GetAttribute("value");
        }

        public void SetDueToInvoiceAnalyticPlan(string value)
        {
            _dueToInvoiceAnalyticPlan = WaitForElementExists(By.Id(DUE_INVOICE_ANALYTIC_PLAN));

            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _dueToInvoiceAnalyticPlan);
            WaitForLoad();
            _dueToInvoiceAnalyticPlan.SetValue(ControlType.TextBox, value);
            WaitPageLoading();
        }

        public void SetSageAutoEnabledForSite(bool sageAutoEnabled)
        {
            _sageAutoEnabled = WaitForElementExists(By.Id(SAGE_AUTO_ENABLED));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _sageAutoEnabled);
            WaitForLoad();

            _sageAutoEnabled = WaitForElementExists(By.Id(SAGE_AUTO_ENABLED));
            _sageAutoEnabled.SetValue(ControlType.CheckBox, sageAutoEnabled);
            // Enregistrement de la donnée (pas de rond busy)
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public bool GetSageAutoEnabledStatusForSite()
        {
            _sageAutoEnabled = WaitForElementExists(By.Id(SAGE_AUTO_ENABLED));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _sageAutoEnabled);
            WaitForLoad();

            _sageAutoEnabled = WaitForElementExists(By.Id(SAGE_AUTO_ENABLED));
            var status = _sageAutoEnabled.Selected;
            WaitForLoad();
            return status;
        }

        public string GetSanitaryAgreement()
        {
            _siteSanitaryAgreement = WaitForElementIsVisible(By.Id(SITE_SANITARY_AGREEMENT));
            return _siteSanitaryAgreement.GetAttribute("value");
        }

        public void SetSanitaryAgreement(string value)
        {
            _siteSanitaryAgreement = WaitForElementIsVisible(By.Id(SITE_SANITARY_AGREEMENT));
            _siteSanitaryAgreement.SetValue(ControlType.TextBox, value);
            WaitPageLoading();
        }

        public string GetEdiGNLCode()
        {
            _ediGNLCode = WaitForElementIsVisible(By.Id(EDI_GLN_CODE));
            return _ediGNLCode.GetAttribute("value");
        }

        public void SetEdiGNLCode(string value)
        {
            _ediGNLCode = WaitForElementIsVisible(By.Id(EDI_GLN_CODE));
            _ediGNLCode.SetValue(ControlType.TextBox, value);
            WaitPageLoading();
        }
        public string GetCountrySelected()
        {
            var country = WaitForElementIsVisible(By.XPath(COUNTRY));
            return country.Text;
        }

        // Sites --> Contacts
        public void ClickToContacts()
        {
            _contactsTab = WaitForElementIsVisible(By.XPath(CONTACTS_SUB_TAB));
            _contactsTab.Click();
            WaitForLoad();
        }

        public void SetInvoiceSageManager(string name, string email)
        {
            _invoiceSageManagerName = WaitForElementExists(By.Id(INVOICE_SAGE_MANAGER_NAME));

            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _invoiceSageManagerName);
            WaitForLoad();
            _invoiceSageManagerName.SetValue(ControlType.TextBox, name);

            _invoiceSageManagerMail = WaitForElementIsVisible(By.Id(INVOICE_SAGE_MANAGER_MAIL));
            _invoiceSageManagerMail.SetValue(ControlType.TextBox, email);

            // Temps d'enregistrement
            WaitPageLoading();
        }
        public void SetInvoiceManager(string name, string email)
        {
            var invoiceManagerName = WaitForElementExists(By.XPath("/html/body/div[2]/div/div/div/div/div[2]/div/div/div/div/div/form/h3[8]"));

            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _invoiceManagerName);
            WaitForLoad();
            _invoiceManagerName.SetValue(ControlType.TextBox, name);

            _invoiceManagerMail = WaitForElementIsVisible(By.XPath(INVOICE_MANAGER_MAIL));
            _invoiceManagerMail.SetValue(ControlType.TextBox, email);

            // Temps d'enregistrement
            WaitPageLoading();
        }

        public string GetInvoiceSageManager()
        {
            _invoiceSageManagerMail = WaitForElementExists(By.Id(INVOICE_SAGE_MANAGER_MAIL));

            return _invoiceSageManagerMail.GetAttribute("value");
        }

        public void SetSupplierInvoiceSageManager(string name, string email)
        {
            _supplierInvoiceSageManagerName = WaitForElementExists(By.Id(SUPPLIER_INVOICE_SAGE_MANAGER_NAME));

            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _supplierInvoiceSageManagerName);
            WaitForLoad();
            _supplierInvoiceSageManagerName.SetValue(ControlType.TextBox, name);

            _supplierInvoiceSageManagerMail = WaitForElementIsVisible(By.Id(SUPPLIER_INVOICE_SAGE_MANAGER_MAIL));
            _supplierInvoiceSageManagerMail.SetValue(ControlType.TextBox, email);

            // Temps d'enregistrement
            WaitPageLoading();
        }

        public string GetSupplierInvoiceSageManager()
        {
            _supplierInvoiceSageManagerMail = WaitForElementExists(By.Id(SUPPLIER_INVOICE_SAGE_MANAGER_MAIL));

            return _supplierInvoiceSageManagerMail.GetAttribute("value");
        }

        // Sites --> Organization
        public void ClickToOrganization()
        {
            _organizationTab = WaitForElementIsVisible(By.XPath(ORGANIZATION_SUB_TAB));
            _organizationTab.Click();
            WaitForLoad();
        }

        // Sites --> Counters
        public void ClickToCounters()
        {
            _counterstab = WaitForElementIsVisible(By.XPath(COUNTERS_TAB));
            _counterstab.Click();
            WaitForLoad();
        }

        public enum PlaceType
        {
            Production,
            Storage,
            OtherSite
        }

        public bool IsOrganizationPresent(string placeTo)
        {
            if (isElementVisible(By.XPath(string.Format(ORGANIZATION_LINE_DEV, placeTo))))
            {

                return true;
            }
            else
            {
                return false;
            }
        }

        public void SearchOrganization(string placeName)
        {
            IWebElement editOrganization;
            editOrganization = WaitForElementExists(By.XPath(string.Format(ORGANIZATION_EDIT_DEV, placeName)));

            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", editOrganization);
            WaitForLoad();

            editOrganization.Click();
            WaitForLoad();
        }

        public void CreateNewOrganization(string placeName)
        {
            _newOrganization = WaitForElementIsVisible(By.XPath(NEW_ORGANIZATION));
            _newOrganization.Click();
            WaitForLoad();

            _organizationName = WaitForElementIsVisible(By.Id(ORGANIZATION_NAME));
            _organizationName.SetValue(ControlType.TextBox, placeName);

            // champ obligatoire (*)
            _contactName = WaitForElementIsVisible(By.Id(CONTACT_NAME));
            _contactName.SetValue(ControlType.TextBox, placeName);


            _saveOrganization = WaitForElementIsVisible(By.XPath(SAVE_ORGANIZATION_DEV));
            _saveOrganization.Click();
            WaitForLoad();
        }
        public void CreateNewOrganization(string placeName, string placeType)
        {
            _newOrganization = WaitForElementIsVisible(By.XPath(NEW_ORGANIZATION));
            _newOrganization.Click();
            WaitForLoad();

            _organizationName = WaitForElementIsVisible(By.Id(ORGANIZATION_NAME));
            _organizationName.SetValue(ControlType.TextBox, placeName);

            _contactName = WaitForElementIsVisible(By.Id(CONTACT_NAME));
            _contactName.SetValue(ControlType.TextBox, placeName);

            switch (placeType.ToLower())
            {
                case "production place":
                    var productionPlace = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div[2]/div/form/div[1]/div[5]/div/input[1]"));
                    productionPlace.Click();
                    break;
                case "storage place":
                    var storagePlace = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div[2]/div/form/div[1]/div[5]/div/input[2]"));
                    storagePlace.Click();
                    break;
                case "other site place":
                    var otherSitePlace = WaitForElementIsVisible(By.XPath("/html/body/div[3]/div/div/div[2]/div/form/div[1]/div[6]/div[1]/input[1]"));
                    otherSitePlace.Click();
                    break;
                case "losses":
                    var losses = WaitForElementIsVisible(By.Id("/html/body/div[3]/div/div/div[2]/div/form/div[1]/div[6]/div[1]/input[2]"));
                    losses.Click();
                    break;
                case "sales":
                    var sales = WaitForElementIsVisible(By.Id("/html/body/div[3]/div/div/div[2]/div/form/div[1]/div[6]/div[1]/input[3]"));
                    sales.Click();
                    break;
              
                  
            }

            _saveOrganization = WaitForElementIsVisible(By.XPath(SAVE_ORGANIZATION_DEV));
            _saveOrganization.Click();
            WaitForLoad();
        }


        public void EditOrganization(PlaceType placeType, string site = null, string placeTo = null, bool isActive = true)
        {
            _isActiveOrganization = WaitForElementExists(By.Id(ORGANIZATION_ACTIVE));
            _isActiveOrganization.SetValue(ControlType.CheckBox, isActive);

            switch (placeType)
            {
                case PlaceType.Production:
                    _productionPlace = WaitForElementIsVisible(By.XPath(ORGA_PROD_PLACE));
                    _productionPlace.SetValue(ControlType.RadioButton, true);
                    break;
                case PlaceType.Storage:
                    _storagePlace = WaitForElementIsVisible(By.XPath(ORGA_STOR_PLACE));
                    _storagePlace.SetValue(ControlType.RadioButton, true);
                    break;
                case PlaceType.OtherSite:
                    _otherSitePlace = WaitForElementIsVisible(By.XPath(ORGA_SITE_PLACE));
                    _otherSitePlace.SetValue(ControlType.RadioButton, true);

                    if (site != null && placeTo != null)
                    {
                        _selectSitePlace = WaitForElementIsVisible(By.Id(SELECT_SITE_PLACE));
                        _selectSitePlace.SetValue(ControlType.DropDownList, site);
                    }

                    if (placeTo != null)
                    {
                        _organizationName = WaitForElementIsVisible(By.Id(ORGANIZATION_NAME));
                        _organizationName.SetValue(ControlType.TextBox, placeTo);

                        // champ obligatoire (*)
                        _contactName = WaitForElementIsVisible(By.Id(CONTACT_NAME));
                        _contactName.SetValue(ControlType.TextBox, placeTo);
                    }

                    break;
            }

            _saveOrganization = WaitForElementIsVisible(By.XPath(SAVE_ORGANIZATION_DEV));
            _saveOrganization.Click();
            WaitForLoad();
        }

        // Fiscal Entities
        public void ClickOnFiscalEntitiesTab()
        {
            _fiscalEntitiesTab = WaitForElementIsVisible(By.Id(FISCAL_ENTITIES_TAB));
            _fiscalEntitiesTab.Click();
            WaitForLoad();
        }

        public ParametersSitesModalPage ClickOnNewFiscalEntityModal()
        {
            //*[@id="tabContentSites"]/a
            // ShowPlusMenu();
            _createNewFiscalEntity = WaitForElementIsVisible(By.XPath(FISCAL_ENTITY_PLUS_NEW));
            _createNewFiscalEntity.Click();
            WaitForLoad();

            return new ParametersSitesModalPage(_webDriver, _testContext);

        }


        private const string NB_LIGNES_FISCAL_ENTITIES = "//*[@id=\"tabContentSites\"]/table/tbody/tr[*]";  // //*[@id="tabContentSites"]/table/tbody/tr[1]
        public int GetLinesFiscalEntitiesCount()
        {
            WaitForLoad();
            return _webDriver.FindElements(By.XPath(NB_LIGNES_FISCAL_ENTITIES)).Count - 1;
        }


        public string SetSageConnectionString(string group, string connectionString)
        {
            _fiscalEntitiesGroupEdit = WaitForElementIsVisible(By.XPath(String.Format(FISCAL_ENTITIES_GROUP_EDIT, group)));
            _fiscalEntitiesGroupEdit.Click();
            WaitForLoad();

            _fiscalEntitySageConnectionString = WaitForElementIsVisible(By.Id(FISCAL_NETITY_SAGE_CONNECTION_STRING));

            if (_fiscalEntitySageConnectionString.GetAttribute("value").Equals(""))
            {
                _fiscalEntitySageConnectionString.SetValue(ControlType.TextBox, connectionString);
            }

            return _fiscalEntitySageConnectionString.GetAttribute("value");
        }

        public bool VerifySageConnectionStringFormat(string connectionString)
        {
            Regex r = new Regex("^SageDbInstance=.+\\|SageDbName=.+", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            System.Text.RegularExpressions.Match m = r.Match(connectionString);

            return m.Success;
        }

        public void UploadLogo(FileInfo fiUpload)
        {
            Assert.IsTrue(fiUpload.Exists, "Fichier d'entrée non trouve");
            var upload = WaitForElementExists(By.Id("FileSent"));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", upload);
            var buttonFile = WaitForElementIsVisible(By.Id("FileSent"));
            buttonFile.SendKeys(fiUpload.FullName);
            //chargement du preview
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public bool CheckIfFirstSiteIsActive()
        {
            var checkbox = WaitForElementExists(By.XPath("/html/body/div[2]/div/div/div/div/div[1]/div/div/form/div/div[2]/table/tbody/tr[2]/td[5]/div/input"));
            return checkbox.Selected;
        }

        public void ActivateFirstSiteInList()
        {
            if (!CheckIfFirstSiteIsActive())
            {
                var checkbox = WaitForElementExists(By.XPath("/html/body/div[2]/div/div/div/div/div[1]/div/div/form/div/div[2]/table/tbody/tr[2]/td[5]/div/input"));
                checkbox.Click();
                WaitForLoad();
            }
        }

        public void DeactivateFirstSiteInList()
        {
            if (CheckIfFirstSiteIsActive())
            {
                var checkbox = WaitForElementExists(By.XPath("/html/body/div[2]/div/div/div/div/div[1]/div/div/form/div/div[2]/table/tbody/tr[2]/td[5]/div/input"));
                checkbox.Click();
                WaitForLoad();
                var btnDeactivate = WaitForElementToBeClickable(By.XPath("/html/body/div[3]/div/div/div/div/form/div[3]/button[3]"));
                btnDeactivate.Click();
                WaitForLoad();
            }
        }
        public string GetFirstSiteName()
        {
            var siteName = WaitForElementExists(By.XPath(FIRST_SITE_NAME)).Text;
            return siteName;
        }
        public string GetPlaceTypeByOrganization(string organization)
        {
            var organisations = _webDriver.FindElements(By.XPath($"//*[@id=\"tabContentSiteDetails\"]/div[2]/div/div[2]/ul/li[*]/div/div/div/div/span[contains(text(),'{organization}')]"));
            var org = organisations.FirstOrDefault();
            return org.Text;
        }

        public string GetSafetyNote()
        {
            var _safetyNote = WaitForElementExists(By.Id(SAFETY_NOTE));
            return _safetyNote.Text;
        }

        public string GetHealthApprovalNumber()
        {
            var sanitaryAgreement = GetSanitaryAgreement();
            var sanitaryDate = WaitForElementExists(By.Id(SANITARY_DATE));
            var healthApprovalNumber = sanitaryAgreement + sanitaryDate.Text;
            return healthApprovalNumber;
        }

        public string GetAirportAgreement()
        {
            var _airoportAgreement = WaitForElementExists(By.XPath(AIRPORT_NUMBER));
            return _airoportAgreement.GetAttribute("value");
        }

        public string GetCustomNumber()
        {
            var customsNumber = WaitForElementExists(By.XPath("//*[@id=\"table-counters\"]/tbody/tr[10]/td[7]/span"));
            return customsNumber.Text;
        }
        public string GetCurrency()
        {
            SelectElement currency = new SelectElement(WaitForElementIsVisible(By.Id(CURRENCY)));
            return currency.SelectedOption.Text;
        }
        public bool CheckFiscalEntityExists(string name, string code, string accounting, int cnt)
        {
            cnt++;
            var col1 = this.WaitForElementExists(By.XPath($"//*[@id=\"tabContentSites\"]/table/tbody/tr[{cnt}]/td[1]"));
            var code_value = col1?.Text;

            var col2 = this.WaitForElementExists(By.XPath($"//*[@id=\"tabContentSites\"]/table/tbody/tr[{cnt}]/td[2]"));
            var name_value = col2?.Text;

            var col3 = this.WaitForElementExists(By.XPath($"//*[@id=\"tabContentSites\"]/table/tbody/tr[{cnt}]/td[3]"));
            var accounting_value = col3?.Text;


            Assert.AreEqual(code, code_value, $"L'entité fiscale avec le code \"{code}\" non trouvée");
            Assert.AreEqual(name, name_value, $"L'entité fiscale avec le nom \"{name}\" non trouvée");
            Assert.AreEqual(accounting, accounting_value, $"L'entité fiscale \"{accounting}\" non trouvée");

            return code.Equals(code_value) && name.Equals(name_value) && accounting.Equals(accounting_value);
        }


        private const string TABLE_SITES_ROW_N_DELETE = "//*[@id=\"tabContentSites\"]/table/tbody/tr[{0}]/td[5]/a[2]";

        public void DeleteFiscalEntityAt(int index)
        {
            index++;
            var deleteBtn = this.WaitForElementToBeClickable(By.XPath(string.Format(TABLE_SITES_ROW_N_DELETE, index)));
            deleteBtn.Click();

            var confirmBtnModal = this.WaitForElementToBeClickable(By.Id("dataConfirmOK"));
            confirmBtnModal.Click();
            WaitForLoad();
        }

        // 1ere row //*[@id="tabContentSites"]/table/tbody/tr[2]/td[1]
        private const string ROW_COLUMN_TEMPLATE = "//*[@id=\"tabContentSites\"]/table/tbody/tr[{0}]/td[{1}]";

        public void DeleteFiscalEntityByValues(string code, string name, string accounting)
        {
            var index = -1;

            var maxi = GetLinesFiscalEntitiesCount();
            bool found = false;

            for (int i = maxi; i >= 0; i--)
            {
                var col1_exists = this.isElementExists(By.XPath(string.Format(ROW_COLUMN_TEMPLATE, i, 1)));
                if (col1_exists)
                {
                    var col1 = this.WaitForElementExists(By.XPath($"//*[@id=\"tabContentSites\"]/table/tbody/tr[{i}]/td[1]"));
                    var code_value = col1?.Text;
                    if (!code_value.Equals(code))
                    {
                        continue;
                    }
                    var col2 = this.WaitForElementExists(By.XPath($"//*[@id=\"tabContentSites\"]/table/tbody/tr[{i}]/td[2]"));
                    var name_value = col2?.Text;
                    if (!name_value.Equals(name))
                    {
                        continue;
                    }
                    var col3 = this.WaitForElementExists(By.XPath($"//*[@id=\"tabContentSites\"]/table/tbody/tr[{i}]/td[3]"));
                    var accounting_value = col3?.Text;
                    if (!accounting_value.Equals(accounting))
                    {
                        continue;
                    }
                    found = true;
                    index = i;
                    break;
                }
            }
            if (found)
            {
                DeleteFiscalEntityAt(index);
            }

        }
        public ParametersSitesModalPage ClickOnNewPurchaseOrdermailbodyModal()
        {
            _purchase_order_mailbody = WaitForElementIsVisible(By.XPath(PURCHASE_ORDER_MAILBODY_PLUS_NEW));
            _purchase_order_mailbody.Click();
            WaitForLoad();

            return new ParametersSitesModalPage(_webDriver, _testContext);

        }
        public void ClickOnPurchaseOrderMailbodyTab()
        {
            _purchase_order_mail_body = WaitForElementIsVisible(By.XPath(PURCHASE_ORDER_MAIL_BODY_TAB));
            _purchase_order_mail_body.Click();
            WaitForLoad();
        }
        public void DeletePurchaseOrderMailbody()
        {
            _delete_purchase_order_mail_body = WaitForElementIsVisible(By.XPath(DELETE_PURCHASE_ORDER_MAILBODY));
            _delete_purchase_order_mail_body.Click();
            _delete_purchase_button = WaitForElementIsVisible(By.Id(DELETE_PURCHASE_BUTTON));
            _delete_purchase_button.Click();

        }
        public string GetClaimManagerMail()
        {
            _claimManagerMail = WaitForElementExists(By.Id(CLAIM_MANAGER_MAIL));

            return _claimManagerMail.GetAttribute("value");
        }

        public string GetCommercialManagerMail()
        {
            _commercialManagerMail = WaitForElementExists(By.Id(COMMERCIAL_MANAGER_MAIL));

            return _commercialManagerMail.GetAttribute("value");
        }

        public void SetCommercialManagerMail(string mail)
        {
            _commercialManagerName = WaitForElementExists(By.Id(COMMERCIAL_MANAGER_NAME));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _commercialManagerName);
            var name = mail.Split('@')[0];
            _commercialManagerName.SetValue(ControlType.TextBox, name);

            _commercialManagerMail = WaitForElementExists(By.Id(COMMERCIAL_MANAGER_MAIL));
            _commercialManagerMail.SetValue(ControlType.TextBox, mail);

            WaitPageLoading();
        }

        public void EnsureCommercialManagerEmail(string currentMail, string newMail)
        {
            if (string.IsNullOrEmpty(currentMail))
            {
                SetCommercialManagerMail(newMail);
            }
        }

        
        public bool IsSiteExisting(string siteName)
        {
            try
            {
                // Chercher le site dans les résultats du filtre
                var siteElement = _webDriver.FindElement(By.XPath($"//table//td[contains(text(), '{siteName}')]"));
                return siteElement != null;
            }
            catch (NoSuchElementException)
            {
                return false; // Le site n'est pas trouvé
            }
        }
        public void DisActivateSite()
        {
            _isActive = WaitForElementExists(By.Id(IS_ACTIVE));
            _isActive.Click();
            WaitForLoad();

            _deactivateSite = WaitForElementIsVisible(By.XPath(DEACTIVATE_SITE));
            _deactivateSite.Click();
            WaitForLoad();

        }
    }
}
