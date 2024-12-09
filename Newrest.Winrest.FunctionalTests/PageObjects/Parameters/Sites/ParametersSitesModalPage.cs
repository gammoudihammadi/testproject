using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Accounting.SupplierInvoices;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Security.Policy;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Sites
{
    public class ParametersSitesModalPage : PageBase
    {

        // _______________________________________ Constantes ___________________________________________

        private const string SITE_NAME = "Name";
        private const string SITE_CODE = "Code";
        private const string SITE_ADDRESS = "Address";
        private const string SITE_ZIP_CODE = "ZipCode";
        private const string SITE_CITY = "City";
        private const string SITE_COUNTRY = "CountryId";

        private const string CREATE = "//*[@id=\"modal-1\"]/div/div/div[2]/div/form/div[2]/button[2]";
        private const string CREATE_DEV = "//*[@id=\"modal-1\"]/div[2]/div/form/div[2]/button[2]";

        private const string CREATE_BUTTON = "//button[text()='Create']";
        private const string SAVE_BUTTON = "last";
        private const string PURCHASE_NAME = "//*[@id=\"ParametersMailBodyModal\"]/div/div/div/div/form/div[2]/div[2]/div/div/div[3]/div[2]";
        private const string test = "//*[@id=\"ParametersMailBodyModal\"]/div/div/div/div/form/div[2]/div[3]";
        private const string FISCAL_ENTITIES = "SelectedFiscalEntities_ms";
        private const string SITES = "SelectedSites_ms";

        // _______________________________________ Variables _____________________________________________

        [FindsBy(How = How.Id, Using = SITE_NAME)]
        private IWebElement _siteName;

        [FindsBy(How = How.Id, Using = SITE_CODE)]
        private IWebElement _siteCode;

        [FindsBy(How = How.Id, Using = SITE_ADDRESS)]
        private IWebElement _siteAddress;

        [FindsBy(How = How.Id, Using = SITE_ZIP_CODE)]
        private IWebElement _siteZIPCode;

        [FindsBy(How = How.Id, Using = SITE_CITY)]
        private IWebElement _siteCity;

        [FindsBy(How = How.Id, Using = SITE_COUNTRY)]
        private IWebElement _siteCountry;

        [FindsBy(How = How.XPath, Using = CREATE)]
        private IWebElement _createBtn;
        [FindsBy(How = How.XPath, Using = PURCHASE_NAME)]
        private IWebElement _purchasename;
        [FindsBy(How = How.XPath, Using = SAVE_BUTTON)]
        private IWebElement _saveBtn;
        [FindsBy(How = How.XPath, Using = FISCAL_ENTITIES)]
        private IWebElement _fiscal_entities;
        [FindsBy(How = How.XPath, Using = SITES)]
        private IWebElement _sites;

        // _________________________________________ Méthodes _______________________________________________

        public ParametersSitesModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        public ParametersSites FillPrincipalField_CreationNewSite(string siteNameCode, string siteAddress, string siteZipCode, string siteCity, string siteCode = null)
        {

            // Définition du nom de site
            _siteName = WaitForElementIsVisible(By.Id(SITE_NAME));
            _siteName.SetValue(ControlType.TextBox, siteNameCode);
            _siteName.SendKeys(Keys.Tab);

            //Définition du code du site
            if (siteCode != null)
            {
                _siteCode = WaitForElementIsVisible(By.Id(SITE_CODE));
                _siteCode.SetValue(ControlType.TextBox, siteCode);
                _siteCode.SendKeys(Keys.Tab);
            }
            else
            {
                _siteCode = WaitForElementIsVisible(By.Id(SITE_CODE));
                _siteCode.SetValue(ControlType.TextBox, siteNameCode);
                _siteCode.SendKeys(Keys.Tab);
            }

            //Définition de l'adresse du site
            _siteAddress = WaitForElementIsVisible(By.Id(SITE_ADDRESS));
            _siteAddress.SetValue(ControlType.TextBox, siteAddress);
            _siteAddress.SendKeys(Keys.Tab);

            //Définition du code postal du site
            _siteZIPCode = WaitForElementIsVisible(By.Id(SITE_ZIP_CODE));
            _siteZIPCode.SetValue(ControlType.TextBox, siteZipCode);
            _siteZIPCode.SendKeys(Keys.Tab);

            //Définition de la ville du site
            _siteCity = WaitForElementIsVisible(By.Id(SITE_CITY));
            _siteCity.SetValue(ControlType.TextBox, siteCity);
            _siteCity.SendKeys(Keys.Tab);

            //Définition du pays du site
            _siteCountry = WaitForElementIsVisible(By.Id(SITE_COUNTRY));
            _siteCountry.SetValue(ControlType.DropDownList, "Spain");
            _siteCountry.SendKeys(Keys.Tab);

            //Validation et création du site
            _createBtn = WaitForElementToBeClickable(By.XPath(CREATE_DEV));
            _createBtn.Click();
            WaitPageLoading();

            return new ParametersSites(_webDriver, _testContext);
        }

        public void CreateNewFiscalEntity(string name, string code, string fiscalEntity, string sageConnectionString, string number, string address, string zipCode, string city, string fiscalEntitySftpFolder)
        {
            // fill mandatory fields (name + code)
            var tbName = this.WaitForElementIsVisible(By.Id("Name"));
            tbName.SetValue(PageBase.ControlType.TextBox, name);

            var tbCode = this.WaitForElementIsVisible(By.Id("Code"));
            tbCode.SetValue(PageBase.ControlType.TextBox, code);

            // select the uyumsoft accounting system
            this.DropdownListSelectById("SelectAccountingSystem", "Uyumsoft");

            // fill other fields
            var tbSageConnectionString = this.WaitForElementIsVisible(By.Id("SageConnectionString"));
            tbSageConnectionString.SetValue(PageBase.ControlType.TextBox, sageConnectionString);

            var tbNumber = this.WaitForElementIsVisible(By.Id("Number"));
            tbNumber.SetValue(PageBase.ControlType.TextBox, number);

            var tbAddress = this.WaitForElementIsVisible(By.Id("Address"));
            tbAddress.SetValue(PageBase.ControlType.TextBox, address);

            var tbZipCode = this.WaitForElementIsVisible(By.Id("ZipCode"));
            tbZipCode.SetValue(PageBase.ControlType.TextBox, zipCode);

            var tbCity = this.WaitForElementIsVisible(By.Id("City"));
            tbCity.SetValue(PageBase.ControlType.TextBox, city);

            var tbFiscalEntitySftpFolder = this.WaitForElementIsVisible(By.Id("FiscalEntitySftpFolder"));
            tbFiscalEntitySftpFolder.SetValue(PageBase.ControlType.TextBox, fiscalEntitySftpFolder);

            // click on create button
            var createBtn = this.WaitForElementToBeClickable(By.XPath(CREATE_BUTTON));
            createBtn.Click();

            WaitPageLoading();

        }
        public void CreatePurchaseOrderMailBody(string fiscalEntity,  string sites, string name)
        {
             _fiscal_entities = this.WaitForElementIsVisible(By.Id(FISCAL_ENTITIES));
            ComboBoxSelectById(new ComboBoxOptions(FISCAL_ENTITIES, fiscalEntity));

             _sites = this.WaitForElementIsVisible(By.Id(SITES));
            ComboBoxSelectById(new ComboBoxOptions(SITES, sites));
            _purchasename = WaitForElementIsVisible(By.XPath(PURCHASE_NAME), nameof(PURCHASE_NAME));
            _purchasename.Click();
            _purchasename.SendKeys(name);

            _saveBtn = this.WaitForElementToBeClickable(By.XPath(test));
            _saveBtn = this.WaitForElementToBeClickable(By.Id(SAVE_BUTTON));
            _saveBtn.Click();

        }
    }
}
