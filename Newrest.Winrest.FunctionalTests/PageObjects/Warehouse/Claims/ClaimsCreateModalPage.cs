using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims
{
    public class ClaimsCreateModalPage : PageBase
    {
        public ClaimsCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes ________________________________________________

        private const string CLAIM_NUMBER = "tb-new-receipt-note-number";
        private const string SITE = "SelectedSiteId";
        private const string SUPPLIER = "drop-down-suppliers";
        private const string DELIVERY_LOCATION = "ClaimNoteDto_SitePlaceId";
        private const string DELIVERY_ORDER_NUMBER = "ClaimNoteDto_DeliveryOrderNumber";
        private const string DATE = "datapicker-new-receipt-note-delivery";      
        private const string ACTIVATED = "ClaimNoteDto_IsActive";       
        private const string CREATE_FROM = "checkBoxCopyFrom";
        private const string SEARCH_NUMBER = "form-copy-from-tbSearchName";
        private const string RECEIPT_NOTES_SELECTED = "//*[@id=\"table-copy-from-rn\"]/tbody/tr[*]/td[2][normalize-space(text())='{0}']/../td[7]/div/input[1]";
        private const string CLAIMS_TAB = "btn-tab-claims";
        private const string SUBMIT = "btn-submit-form-create-receipt-note";
        private const string VALIDATION_ERROR = "//*/span[contains(@class,'text-danger') and contains(@class,'field-validation-error')]";
        private const string ID_CLAIM = "/html/body/div[3]/div/div/div/div[2]/div/form/div/div[1]/div[1]/div/div/input";

        //__________________________________ Variables _________________________________________________

        [FindsBy(How = How.Id, Using = CLAIM_NUMBER)]
        private IWebElement _claimNumber;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = SUPPLIER)]
        private IWebElement _supplier;

        [FindsBy(How = How.Id, Using = DELIVERY_LOCATION)]
        private IWebElement _deliveryLocation;

        [FindsBy(How = How.Id, Using = DELIVERY_ORDER_NUMBER)]
        private IWebElement _deliveryOrderNumber;

        [FindsBy(How = How.Id, Using = DATE)]
        private IWebElement _date;

        [FindsBy(How = How.Id, Using = ACTIVATED)]
        private IWebElement _activated;

        [FindsBy(How = How.Id, Using = CREATE_FROM)]
        private IWebElement _createFrom;

        [FindsBy(How = How.Id, Using = SEARCH_NUMBER)]
        private IWebElement _searchNumber;
        
        [FindsBy(How = How.XPath, Using = RECEIPT_NOTES_SELECTED)]
        private IWebElement _receiptNoteSelected;

        [FindsBy(How = How.Id, Using = CLAIMS_TAB)]
        private IWebElement _claimTab;

        [FindsBy(How = How.Id, Using = SUBMIT)]
        private IWebElement _submit;
        
        


        //___________________________________________ Méthodes ______________________________________________________

        public void FillField_CreatNewClaims(DateTime date, string site, string supplier, string placeTo, bool isActivate = true, string createFrom = null, string numberToCopy = null)
        {
            Random rnd = new Random();

            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _supplier = WaitForElementIsVisible(By.Id(SUPPLIER));
            _supplier.SetValue(ControlType.DropDownList, supplier);

            _deliveryLocation = WaitForElementIsVisible(By.Id(DELIVERY_LOCATION));
            _deliveryLocation.SetValue(ControlType.DropDownList, placeTo);

            _deliveryOrderNumber = WaitForElementIsVisible(By.Id(DELIVERY_ORDER_NUMBER));
            _deliveryOrderNumber.SetValue(ControlType.TextBox, rnd.Next().ToString());

            _activated = WaitForElementExists(By.Id(ACTIVATED));
            _activated.SetValue(ControlType.CheckBox, isActivate);

            _date = WaitForElementIsVisible(By.Id(DATE));
            _date.SetValue(ControlType.DateTime, date);
            _date.SendKeys(Keys.Tab);

            if (createFrom == "Receipt Note")
            {
                if (numberToCopy != null)
                {
                    _searchNumber = WaitForElementIsVisible(By.Id(SEARCH_NUMBER));
                    _searchNumber.SetValue(ControlType.TextBox, numberToCopy);
                    WaitForLoad();

                    _receiptNoteSelected = WaitForElementExists(By.XPath(string.Format(RECEIPT_NOTES_SELECTED, numberToCopy)));
                    _receiptNoteSelected.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    WaitPageLoading();
                }
            }
            else if (createFrom == "Claim")
            {
                //tab Claims
                _claimTab = WaitForElementIsVisible(By.Id(CLAIMS_TAB));
                _claimTab.Click();

                _searchNumber = WaitForElementIsVisible(By.Id(SEARCH_NUMBER));
                _searchNumber.SetValue(ControlType.TextBox, numberToCopy);
                WaitForLoad();

                var claimSelected = WaitForElementExists(By.XPath(string.Format(RECEIPT_NOTES_SELECTED, numberToCopy)));
                claimSelected.SetValue(ControlType.CheckBox, true);
                WaitForLoad();
            }
            else if (createFrom == "Purchase Order")
            {
                //tab Claims
                _claimTab = WaitForElementIsVisible(By.Id("btn-tab-purchase-order-claim"));
                _claimTab.Click();
                WaitForLoad();

                _searchNumber = WaitForElementIsVisible(By.Id(SEARCH_NUMBER));
                _searchNumber.SetValue(ControlType.TextBox, numberToCopy);
                WaitForLoad();

                var claimSelected = WaitForElementExists(By.Id("item_IsSelected"));
                claimSelected.Click();
                WaitForLoad();
            }
            else if (createFrom == "Supplier Invoice")
            {
                // combobox "Created from SI"
                var createdFromSI = WaitForElementIsVisible(By.Id("ddlCreatedFrom"));
                SelectElement select = new SelectElement(createdFromSI);
                select.SelectByText("Created from SI");
                // liste complète avant d'être filtré
                WaitForLoad();

                // filtrage
                _searchNumber = WaitForElementIsVisible(By.Id(SEARCH_NUMBER));
                _searchNumber.SetValue(ControlType.TextBox, numberToCopy);
                WaitForLoad();
                WaitPageLoading();

                var claimSelected = WaitForElementExists(By.XPath("//*/td[contains(text(),'" + numberToCopy + "')]/../td[7]/div/input[1]"));
                claimSelected.SetValue(ControlType.CheckBox, true);
                WaitForLoad();
                WaitPageLoading();
            }
            else
            {
                _createFrom = WaitForElementExists(By.Id(CREATE_FROM));
                _createFrom.SetValue(ControlType.CheckBox, false);
                WaitForLoad();
            }
        }


        public ClaimsItem Submit()
        {
            _submit = WaitForElementIsVisible(By.Id(SUBMIT));
            new Actions(_webDriver).MoveToElement(_submit).Click().Perform();
            WaitPageLoading();
            WaitForLoad();
            return new ClaimsItem(_webDriver, _testContext);
        }

        public string GetClaimId()
        {
            _claimNumber = WaitForElementIsVisible(By.Id(CLAIM_NUMBER));
            return _claimNumber.GetAttribute("value");
        }

        public string ValidationError()
        {
            var _validationError = WaitForElementIsVisible(By.XPath(VALIDATION_ERROR));
            return _validationError.Text;
        }

        public void FillFieldReceiptNoteSearch(string rnNumber)
        {
            var filtre = WaitForElementIsVisible(By.Id("form-copy-from-tbSearchName"));
            filtre.SendKeys(rnNumber);
            filtre.SendKeys(Keys.Enter);
            WaitForLoad();
        }
        public string GetIdClaim()
        {

            var element = WaitForElementIsVisible(By.XPath(ID_CLAIM));
            return element.GetAttribute("value");
        }
    }
}
