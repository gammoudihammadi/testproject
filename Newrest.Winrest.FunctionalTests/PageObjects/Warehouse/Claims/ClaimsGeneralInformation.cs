using DocumentFormat.OpenXml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Claims
{
    public class ClaimsGeneralInformation : PageBase
    {
        public ClaimsGeneralInformation(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //_________________________________ Constantes ______________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        // Informations
        private const string CLAIM_NUMBER = "claimnote-number";
        private const string CLAIM_DELIVERY_ORDER_NUMBER = "ClaimNote_DeliveryOrderNumber";
        private const string STATUS = "ClaimNote_Status";
        private const string CLAIM_COMMENT = "ClaimNote_Comment";
        private const string CLAIM_SITE = "ClaimNote_SitePlace_Site_Name";
        private const string CLAIM_SUPPLIER = "ClaimNote_Supplier_Name";
        private const string CLAIM_SUPPLIER_INVOICE = "//*[@id='form-create-claim-note']/div/div[4]/div[2]/div/div/ul/li/a";
        private const string CLAIM_DELIVERY_LOCATION = "ClaimNote_SitePlace_Title";
        private const string CLAIM_RELATED_RN_DELIVERY_LOCATION = "//*/label[text()='Related receipt note delivery order number : ']/../div";
        private const string CLAIM_STATUS = "ClaimNote_Status";
        private const string CLAIMS_ITEM = "hrefTabContentItems";
        private const string CLAIMS_DATE = "//*[@id=\"form-create-claim-note\"]/div/div[4]/div[1]/div/div/input[1]";

        //_________________________________ Variables _______________________________________

        // General
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        // Informations
        [FindsBy(How = How.Id, Using = CLAIM_NUMBER)]
        private IWebElement _claimNumber;

        [FindsBy(How = How.Id, Using = CLAIM_DELIVERY_ORDER_NUMBER)]
        private IWebElement _claimDeliveryOrderNumber;

        [FindsBy(How = How.Id, Using = STATUS)]
        private IWebElement _status;

        [FindsBy(How = How.Id, Using = CLAIM_COMMENT)]
        private IWebElement _claimComment;

        [FindsBy(How = How.Id, Using = CLAIM_SITE)]
        private IWebElement _claimSite;

        [FindsBy(How = How.Id, Using = CLAIM_SUPPLIER)]
        private IWebElement _claimSupplier;

        [FindsBy(How = How.XPath, Using = CLAIM_SUPPLIER_INVOICE)]
        private IWebElement _claimSupplierInvoice;

        [FindsBy(How = How.Id, Using = CLAIM_DELIVERY_LOCATION)]
        private IWebElement _claimDeliveryLocation;

        [FindsBy(How = How.Id, Using = CLAIM_STATUS)]
        private IWebElement _claimStatus;

        [FindsBy(How = How.Id, Using = CLAIMS_ITEM)]
        private IWebElement _claimsItem;

        [FindsBy(How = How.XPath, Using = CLAIMS_DATE)]
        private IWebElement _claimsDate;

        //________________________________ Méthodes ________________________________________

        // General
        public ClaimsPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new ClaimsPage(_webDriver, _testContext);
        }

        // Informations
        public void SetStatus(string value)
        {
            _status = WaitForElementIsVisible(By.Id(STATUS));
            _status.SetValue(ControlType.DropDownList, value);

            // Ophélie : temps de prise en compte de la donnée
            Thread.Sleep(1000);
        }

        public string GetClaimNumber()
        {
            _claimNumber = WaitForElementIsVisibleNew(By.Id(CLAIM_NUMBER));
            return _claimNumber.GetAttribute("value");
        }

        public string GetDeliveryOrderNumber()
        {
            _claimDeliveryOrderNumber = WaitForElementIsVisible(By.Id(CLAIM_DELIVERY_ORDER_NUMBER));
            return _claimDeliveryOrderNumber.GetAttribute("value");
        }

        public string GetClaimComment()
        {
            _claimComment = WaitForElementIsVisible(By.Id(CLAIM_COMMENT));
            return _claimComment.GetAttribute("value");
        }

        public string GetClaimSite()
        {
            _claimSite = WaitForElementIsVisible(By.Id(CLAIM_SITE));
            return _claimSite.GetAttribute("value");
        }

        public string GetClaimDate()
        {
            _claimsDate = WaitForElementIsVisible(By.XPath(CLAIMS_DATE));
            string dateTimeValue = _claimsDate.GetAttribute("value");
            return dateTimeValue.Split(' ')[0];
        }

        public string GetClaimSupplier()
        {
            _claimSupplier = WaitForElementIsVisible(By.Id(CLAIM_SUPPLIER));
            return _claimSupplier.GetAttribute("value");
        }

        public string GetClaimDeliveryLocation()
        {
            _claimDeliveryLocation = WaitForElementIsVisible(By.Id(CLAIM_DELIVERY_LOCATION));
            return _claimDeliveryLocation.GetAttribute("value");
        }

        public string GetStatus()
        {
            _claimStatus = WaitForElementIsVisible(By.Id(CLAIM_STATUS));
            return _claimStatus.GetAttribute("value");
        }

        public ClaimsItem ClickOnItems()
        {
            _claimsItem = WaitForElementIsVisible(By.Id(CLAIMS_ITEM));
            _claimsItem.Click();
            WaitForLoad();

            return new ClaimsItem(_webDriver, _testContext);
        }

        public string GetClaimSupplierInvoice()
        {
            _claimSupplierInvoice = WaitForElementIsVisible(By.XPath(CLAIM_SUPPLIER_INVOICE));
            return _claimSupplierInvoice.Text.Split(' ')[0];
        }

        public string GetRelatedReceiptNoteDeliveryOrderNumber()
        {
            var rnDelivery = WaitForElementIsVisible(By.XPath(CLAIM_RELATED_RN_DELIVERY_LOCATION));
            return rnDelivery.Text;
        }

        public void SetDeliveryOrderNumber(string deliveryOrderNumber)
        {
            var claimDeliveryOrderNumber = WaitForElementIsVisible(By.Id(CLAIM_DELIVERY_ORDER_NUMBER));
            claimDeliveryOrderNumber.SetValue(ControlType.TextBox, deliveryOrderNumber);
        }

        public void SetComment(string comment)
        {
            var _claimComment = WaitForElementExists(By.Id(CLAIM_COMMENT));
            new Actions(_webDriver).MoveToElement(_claimComment).Perform();
            _claimComment = WaitForElementIsVisible(By.Id(CLAIM_COMMENT));
            // TextArea
            _claimComment.Clear();
            _claimComment.SendKeys(comment);
            WaitForLoad();
        }

        public void PageUp()
        {
            WaitForLoad();
            Thread.Sleep(2000);
            WaitForLoad();
            _webDriver.Navigate().Refresh();
        }

        public string GetRelatedReceiptNoteNumber()
        {
            var rnNumber = WaitForElementIsVisible(By.XPath("//*[@id=\"form-create-claim-note\"]/div/div[8]/div[1]/div/div/a"));
            return rnNumber.Text;
        }
        public string GetRelatedReceiptNoteNumberLink()
        {
            var rnNumber = WaitForElementIsVisible(By.XPath("//*[@id=\"form-create-claim-note\"]/div/div[8]/div[1]/div/div/a"));
            return rnNumber.GetAttribute("href");
        }
    }
}
