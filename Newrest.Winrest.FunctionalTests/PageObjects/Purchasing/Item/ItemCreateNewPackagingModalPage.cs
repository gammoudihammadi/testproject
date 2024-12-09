using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System.Security.Cryptography.Xml;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Item
{
    public class ItemCreateNewPackagingModalPage : PageBase
    {

        public ItemCreateNewPackagingModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _____________________________________________ Constantes __________________________________________

        private const string PACKAGING_SITE = "NewPackaging_SiteId";
        private const string PACKAGING_SITE_SELECT = "/html/body/div[4]/div/div/div/div/form/div[2]/div[1]/div/select/option[text()= \"{0}\"]";
        private const string PACKAGING_NAME = "PackagingUnitListVM_ItemUnitPackagingTypeId";
        private const string PACKAGING_NAME_SELECT = "/html/body/div[4]/div/div/div/div/form/div[2]/div[2]/div/select/option[text() = \"{0}\"]";
        private const string PACKAGING_STORAGE_QUANTITY = "NewPackaging_ItemUnitStorageTypeValue";
        private const string PACKAGING_STORAGE_UNIT = "NewPackaging_ItemUnitStorageTypeId";
        private const string PACKAGING_STORAGE_UNIT_SELECT = "/html/body/div[4]/div/div/div/div/form/div[2]/div[4]/div/select/option[text()=\"{0}\"]";
        private const string PACKAGING_QUANTITY = "NewPackaging_ItemUnitProductionQuantity";
        private const string PACKAGING_SUPPLIER = "NewPackaging_SupplierId";
        private const string PACKAGING_SUPPLIER_SELECT = "/html/body/div[4]/div/div/div/div/form/div[2]/div[6]/div/select/option[text()=\"{0}\"]";
        private const string PACKAGING_SUPPLIER_REF = "NewPackaging_RefSupplier";
        private const string PACKAGING_UNIT_PRICE = "NewPackaging_UnitPrice";
        private const string CREATE = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[3]/button[2]";
        private const string PACKAGING_LIMIT_QUANTITY = "//*[@id=\"NewPackaging_MinimumAlertThreshold\"]";
        private const string SITE_ERROR = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[1]/div/span/span";
        private const string PACKAGING_ERROR = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[2]/div/span/span";
        private const string QUANTITY_ERROR = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[3]/div/span/span";
        private const string STORAGE_UNIT_ERROR = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[4]/div/span[1]/span";
        private const string SUPPLIER_ERROR = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[6]/div/span/span";
        private const string YIELD = "//*[@id=\"NewPackaging_Loss\"]";
        private const string ROUNDING_TO_SU = "//*[@id=\"NewPackaging_IsRounded\"]";
        private const string REFERENCE = "//*[@id=\"Reference\"]";
        private const string CHECKALL = "/html/body/div[12]/div/ul/li[1]/a/span[2]";
        private const string SITE_ID = "SelectedSites_ms";


        // _____________________________________________ Variables ___________________________________________

        [FindsBy(How = How.Id, Using = PACKAGING_SITE)]
        private IWebElement _packagingSite;

        [FindsBy(How = How.Id, Using = PACKAGING_NAME)]
        private IWebElement _packagingName;

        [FindsBy(How = How.Id, Using = PACKAGING_STORAGE_QUANTITY)]
        private IWebElement _packagingStorageQty;

        [FindsBy(How = How.Id, Using = PACKAGING_STORAGE_UNIT)]
        private IWebElement _packagingStorageUnit;

        [FindsBy(How = How.Id, Using = PACKAGING_QUANTITY)]
        private IWebElement _packagingQty;

        [FindsBy(How = How.Id, Using = PACKAGING_SUPPLIER)]
        private IWebElement _packagingSupplier;

        [FindsBy(How = How.Id, Using = PACKAGING_SUPPLIER_REF)]
        private IWebElement _packagingSupplierRef;

        [FindsBy(How = How.Id, Using = PACKAGING_UNIT_PRICE)]
        private IWebElement _unitPrice;

        [FindsBy(How = How.XPath, Using = CREATE)]
        private IWebElement _createPackagingBtn;

        [FindsBy(How = How.XPath, Using = YIELD)]
        private IWebElement _yield;

        [FindsBy(How = How.XPath, Using = ROUNDING_TO_SU)]
        private IWebElement _roundingToSu;
        // Message erreur
        [FindsBy(How = How.XPath, Using = SITE_ERROR)]
        private IWebElement _siteErrorMessage;

        [FindsBy(How = How.XPath, Using = PACKAGING_ERROR)]
        private IWebElement _packagingErrorMessage;

        [FindsBy(How = How.XPath, Using = QUANTITY_ERROR)]
        private IWebElement _quantityErrorMessage;

        [FindsBy(How = How.XPath, Using = STORAGE_UNIT_ERROR)]
        private IWebElement _storageUnitErrorMessage;

        [FindsBy(How = How.XPath, Using = SUPPLIER_ERROR)]
        private IWebElement _supplierErrorMessage;

        [FindsBy(How = How.XPath, Using = REFERENCE)]
        private IWebElement _reference;

        [FindsBy(How = How.Id, Using = CHECKALL)]
        private IWebElement _checkall;

        [FindsBy(How = How.Id, Using = SITE_ID)]
        private IWebElement _siteid;
        // _____________________________________________ Méthodes ____________________________________________

        public ItemGeneralInformationPage FillField_CreateNewPackaging(string site, string packaging, string storageQty,
            string storageUnit, string qty, string supplier, string limitQty = null, string supplierRef = null, string unitPrice = "10", bool roundingToSu = true)
        {
            return FillField_CreateNewPackagingMultiSites(new string[] { site }, packaging, storageQty, storageUnit, qty, supplier, limitQty, supplierRef, unitPrice, roundingToSu);
        }

        public ItemGeneralInformationPage FillField_CreateNewPackagingMultiSites(string[] sites, string packaging, string storageQty,
            string storageUnit, string qty, string supplier, string limitQty = null, string supplierRef = null, string unitPrice = "10",bool roundingToSu = true)
        {
            var firstSite = true;
            foreach (var site in sites)
            {
                ComboBoxOptions cbOpt = new ComboBoxOptions("SelectedSites_ms", site.ToString(), false) { ClickCheckAllAtStart = false, ClickUncheckAllAtStart = firstSite };
                ComboBoxSelectById(cbOpt);
                firstSite = false;
            }

            _packagingName = WaitForElementIsVisible(By.Id(PACKAGING_NAME));
            _packagingName.Click();


            if (isElementVisible(By.XPath(string.Format(PACKAGING_NAME_SELECT, packaging))))
            {
                var packagingName = WaitForElementIsVisible(By.XPath(string.Format(PACKAGING_NAME_SELECT, packaging)));
                packagingName.Click();
                _packagingName.SendKeys(Keys.Escape);
            }
            else
            {
                _packagingName.SetValue(ControlType.DropDownList, packaging);
            }

            _packagingStorageQty = WaitForElementIsVisible(By.Id(PACKAGING_STORAGE_QUANTITY));
            _packagingStorageQty.SetValue(ControlType.TextBox, storageQty);

            _packagingStorageUnit = WaitForElementIsVisible(By.Id("PackagingUnitListVM_ItemUnitStorageTypeId"));
            _packagingStorageUnit.Click();

            if (isElementVisible(By.XPath(string.Format(PACKAGING_STORAGE_UNIT_SELECT, storageUnit))))
            {
                var packagingStorageUnit = WaitForElementIsVisible(By.XPath(string.Format(PACKAGING_STORAGE_UNIT_SELECT, storageUnit)));
                packagingStorageUnit.Click();
                _packagingStorageQty.SendKeys(Keys.Tab);
            }
            else
            {
                _packagingStorageUnit.SetValue(ControlType.DropDownList, storageUnit);
            }

            _packagingQty = WaitForElementIsVisible(By.Id(PACKAGING_QUANTITY));
            _packagingQty.SetValue(ControlType.TextBox, qty);

            _packagingSupplier = WaitForElementIsVisible(By.Id(PACKAGING_SUPPLIER));
            _packagingSupplier.Click();
            if (isElementVisible(By.XPath(string.Format(PACKAGING_SUPPLIER_SELECT, supplier))))
            {
                var packagingSite = WaitForElementIsVisible(By.XPath(string.Format(PACKAGING_SUPPLIER_SELECT, supplier)));
                packagingSite.Click();
                _packagingSupplier.SendKeys(Keys.Escape);
            }
            else
            {
                _packagingSupplier.SetValue(ControlType.DropDownList, supplier);
            }

            if (supplierRef != null)
            {
                _packagingSupplierRef = WaitForElementIsVisible(By.Id(PACKAGING_SUPPLIER_REF));
                _packagingSupplierRef.SetValue(ControlType.TextBox, supplierRef);
            }

            _unitPrice = WaitForElementIsVisible(By.Id(PACKAGING_UNIT_PRICE));
            _unitPrice.SetValue(ControlType.TextBox, unitPrice);

            if(limitQty != null)
            {
                var _limitQty = WaitForElementIsVisible(By.XPath(PACKAGING_LIMIT_QUANTITY));
                _limitQty.SetValue(ControlType.TextBox, limitQty);
            }
            _roundingToSu = WaitForElementExists(By.XPath(ROUNDING_TO_SU));
            _roundingToSu.SetValue(ControlType.CheckBox, roundingToSu);



            WaitForLoad();
            _createPackagingBtn = WaitForElementToBeClickable(By.XPath("//*[@id='modal-1']/div/div/form/div[3]/button[2]"));
            _createPackagingBtn.Click();
            
            WaitPageLoading();
            WaitForLoad();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }
        public void Checkall()
        {
            _checkall = WaitForElementIsVisible(By.XPath(CHECKALL));
            _checkall.Click();

        }

        public ItemGeneralInformationPage FillField_CreateNewPackagingAllSites( string packaging, string storageQty,
           string storageUnit, string qty, string supplier, string limitQty = null, string supplierRef = null, string unitPrice = "10", bool roundingToSu = true)
        {
            _siteid = WaitForElementIsVisible(By.Id(SITE_ID));
            _siteid.Click();
            Checkall();
            _siteid.Click();

            _packagingName = WaitForElementIsVisible(By.Id(PACKAGING_NAME));
            _packagingName.Click();


            if (isElementVisible(By.XPath(string.Format(PACKAGING_NAME_SELECT, packaging))))
            {
                var packagingName = WaitForElementIsVisible(By.XPath(string.Format(PACKAGING_NAME_SELECT, packaging)));
                packagingName.Click();
                _packagingName.SendKeys(Keys.Escape);
            }
            else
            {
                _packagingName.SetValue(ControlType.DropDownList, packaging);
            }

            _packagingStorageQty = WaitForElementIsVisible(By.Id(PACKAGING_STORAGE_QUANTITY));
            _packagingStorageQty.SetValue(ControlType.TextBox, storageQty);

            _packagingStorageUnit = WaitForElementIsVisible(By.Id("PackagingUnitListVM_ItemUnitStorageTypeId"));
            _packagingStorageUnit.Click();

            if (isElementVisible(By.XPath(string.Format(PACKAGING_STORAGE_UNIT_SELECT, storageUnit))))
            {
                var packagingStorageUnit = WaitForElementIsVisible(By.XPath(string.Format(PACKAGING_STORAGE_UNIT_SELECT, storageUnit)));
                packagingStorageUnit.Click();
                _packagingStorageQty.SendKeys(Keys.Tab);
            }
            else
            {
                _packagingStorageUnit.SetValue(ControlType.DropDownList, storageUnit);
            }

            _packagingQty = WaitForElementIsVisible(By.Id(PACKAGING_QUANTITY));
            _packagingQty.SetValue(ControlType.TextBox, qty);

            _packagingSupplier = WaitForElementIsVisible(By.Id(PACKAGING_SUPPLIER));
            _packagingSupplier.Click();
            if (isElementVisible(By.XPath(string.Format(PACKAGING_SUPPLIER_SELECT, supplier))))
            {
                var packagingSite = WaitForElementIsVisible(By.XPath(string.Format(PACKAGING_SUPPLIER_SELECT, supplier)));
                packagingSite.Click();
                _packagingSupplier.SendKeys(Keys.Escape);
            }
            else
            {
                _packagingSupplier.SetValue(ControlType.DropDownList, supplier);
            }

            if (supplierRef != null)
            {
                _packagingSupplierRef = WaitForElementIsVisible(By.Id(PACKAGING_SUPPLIER_REF));
                _packagingSupplierRef.SetValue(ControlType.TextBox, supplierRef);
            }

            _unitPrice = WaitForElementIsVisible(By.Id(PACKAGING_UNIT_PRICE));
            _unitPrice.SetValue(ControlType.TextBox, unitPrice);

            if (limitQty != null)
            {
                var _limitQty = WaitForElementIsVisible(By.XPath(PACKAGING_LIMIT_QUANTITY));
                _limitQty.SetValue(ControlType.TextBox, limitQty);
            }
            _roundingToSu = WaitForElementExists(By.XPath(ROUNDING_TO_SU));
            _roundingToSu.SetValue(ControlType.CheckBox, roundingToSu);



            WaitForLoad();
            _createPackagingBtn = WaitForElementToBeClickable(By.XPath("//*[@id='modal-1']/div/div/form/div[3]/button[2]"));
            _createPackagingBtn.Click();

            WaitPageLoading();
            WaitForLoad();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

        public void setYield(string yield)
        {
            _yield = WaitForElementIsVisible(By.XPath(YIELD));
            _yield.SetValue(ControlType.TextBox, yield);
        }

        public void FillFieldReference(string ArticleCode)
        {
            WaitPageLoading();
            _reference = WaitForElementIsVisible(By.XPath(REFERENCE));
            _reference.SetValue(ControlType.TextBox, ArticleCode);
        }
        
        public ItemGeneralInformationPage NoFillFieldPackaging()
        {
            // Click sur le bouton "Create"
            _createPackagingBtn = WaitForElementExists(By.XPath("//*/button[@value='Create']"));
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            javaScriptExecutor.ExecuteScript("arguments[0].scrollIntoView(true);", _createPackagingBtn);
            WaitForLoad();

            _createPackagingBtn.Click();
            WaitForLoad();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

        // Message d'erreur
        public bool ErrorMessageSiteRequired()
        {
            _siteErrorMessage = WaitForElementToBeClickable(By.Id("siteserror"));

            string expectedMessage = "At least one site must be selected";
            if (_siteErrorMessage.Text == expectedMessage)
            {
                return true;
            }

            return false;
        }

        public bool ErrorMessagePackagingRequired()
        {
            _packagingErrorMessage = WaitForElementToBeClickable(By.Id("packagingserror"));

            string expectedMessage = "This field is required";
            if (_packagingErrorMessage.Text == expectedMessage)
            {
                return true;
            }

            return false;
        }

        public bool ErrorMessageQuantityRequired()
        {

            _quantityErrorMessage = WaitForElementToBeClickable(By.XPath("//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[3]/div/span"));

            string expectedMessage = "error_range_greater_than_0.0";
            if (_quantityErrorMessage.Text == expectedMessage)
            {
                return true;
            }

            return false;
        }

        public bool ErrorMessageStorageUnitRequired()
        {
            _storageUnitErrorMessage = WaitForElementToBeClickable(By.Id("storageerror"));

            string expectedMessage = "This field is required";
            if (_storageUnitErrorMessage.Text == expectedMessage)
            {
                return true;
            }

            return false;
        }
        public bool ErrorMessageSupplierRequired()
        {
            _supplierErrorMessage = WaitForElementToBeClickable(By.Id("supplierIderror"));

            string expectedMessage = "This field is required";
            if (_supplierErrorMessage.Text == expectedMessage)
            {
                return true;
            }

            return false;
        }
        public ItemGeneralInformationPage FillField_CreateNewPackagingWithoutPrice(string site, string packaging, string storageQty,
           string storageUnit, string qty, string supplier)
        {
            return FillField_CreateNewPackagingMultiSitesWithoutPrice(new string[] { site }, packaging, storageQty, storageUnit, qty, supplier);
        }

        public ItemGeneralInformationPage FillField_CreateNewPackagingMultiSitesWithoutPrice(string[] sites, string packaging, string storageQty,
            string storageUnit, string qty, string supplier )
        {
            var firstSite = true;
            foreach (var site in sites)
            {
                ComboBoxOptions cbOpt = new ComboBoxOptions("SelectedSites_ms", site.ToString(), false) { ClickCheckAllAtStart = false, ClickUncheckAllAtStart = firstSite };
                ComboBoxSelectById(cbOpt);
                firstSite = false;
            }

            _packagingName = WaitForElementIsVisible(By.Id(PACKAGING_NAME));
            _packagingName.Click();


            if (isElementVisible(By.XPath(string.Format(PACKAGING_NAME_SELECT, packaging))))
            {
                var packagingName = WaitForElementIsVisible(By.XPath(string.Format(PACKAGING_NAME_SELECT, packaging)));
                packagingName.Click();
                _packagingName.SendKeys(Keys.Escape);
            }
            else
            {
                _packagingName.SetValue(ControlType.DropDownList, packaging);
            }

            _packagingStorageQty = WaitForElementIsVisible(By.Id(PACKAGING_STORAGE_QUANTITY));
            _packagingStorageQty.SetValue(ControlType.TextBox, storageQty);

            _packagingStorageUnit = WaitForElementIsVisible(By.Id("PackagingUnitListVM_ItemUnitStorageTypeId"));
            _packagingStorageUnit.Click();

            if (isElementVisible(By.XPath(string.Format(PACKAGING_STORAGE_UNIT_SELECT, storageUnit))))
            {
                var packagingStorageUnit = WaitForElementIsVisible(By.XPath(string.Format(PACKAGING_STORAGE_UNIT_SELECT, storageUnit)));
                packagingStorageUnit.Click();
                _packagingStorageQty.SendKeys(Keys.Tab);
            }
            else
            {
                _packagingStorageUnit.SetValue(ControlType.DropDownList, storageUnit);
            }

            _packagingQty = WaitForElementIsVisible(By.Id(PACKAGING_QUANTITY));
            _packagingQty.SetValue(ControlType.TextBox, qty);

            _packagingSupplier = WaitForElementIsVisible(By.Id(PACKAGING_SUPPLIER));
            _packagingSupplier.Click();
            if (isElementVisible(By.XPath(string.Format(PACKAGING_SUPPLIER_SELECT, supplier))))
            {
                var packagingSite = WaitForElementIsVisible(By.XPath(string.Format(PACKAGING_SUPPLIER_SELECT, supplier)));
                packagingSite.Click();
                _packagingSupplier.SendKeys(Keys.Escape);
            }
            else
            {
                _packagingSupplier.SetValue(ControlType.DropDownList, supplier);
            }

            WaitForLoad();
            _createPackagingBtn = WaitForElementToBeClickable(By.XPath("//*[@id='modal-1']/div/div/form/div[3]/button[2]"));
            _createPackagingBtn.Click();

            WaitPageLoading();
            WaitForLoad();

            return new ItemGeneralInformationPage(_webDriver, _testContext);
        }

    }
}
