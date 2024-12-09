using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;


namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes
{
    public class ReceiptNotesCreateModalPage : PageBase
    {

        public ReceiptNotesCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //_________________________________ Constantes ___________________________________________________

        private const string SITE = "SelectedSiteId";
        private const string SUPPLIER = "drop-down-suppliers";
        private const string DELIVERY_LOCATION = "SelectedSitePlaceId";
        private const string DELIVERY_ORDER_NUMBER = "ReceiptNote_DeliveryOrderNumber";
        private const string DELIVERY_DATE = "datapicker-new-receipt-note-delivery";
        private const string ACTIVATED = "ReceiptNote_IsActive";
        private const string CREATE_FROM = "checkBoxCopyFrom";
        private const string COMMENT = "ReceiptNote_Comment";
        private const string PURCHASE_ORDER_TAB = "btn-tab-purchase-order";
        private const string PURCHASE_ORDER_NUMBER = "form-copy-from-tbSearchName";        
        private const string PURCHASE_ORDER_DATE = "copy-from-date-picker-start";
        private const string PURCHASE_ORDER_SELECTED = "//*[@id=\"table-copy-from-po\"]/tbody/tr[*]/td[2][normalize-space(text()) = '{0}']/../td[6]/div/input";

        private const string SUBMIT = "btn-submit-form-create-receipt-note";

        private const string VALIDATOR_BACKDATING = "You are backdating your document, this is not allowed. Contact your Winrest Referent";

        //________________________________ Variables ____________________________________________________

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = SUPPLIER)]
        private IWebElement _supplier;

        [FindsBy(How = How.Id, Using = DELIVERY_LOCATION)]
        private IWebElement _deliveryLocation;

        [FindsBy(How = How.Id, Using = DELIVERY_ORDER_NUMBER)]
        private IWebElement _deliveryOrderNumber;

        [FindsBy(How = How.Id, Using = DELIVERY_DATE)]
        private IWebElement _deliveryDate;

        [FindsBy(How = How.Id, Using = ACTIVATED)]
        private IWebElement _activated;

        [FindsBy(How = How.Id, Using = CREATE_FROM)]
        private IWebElement _createFrom;

        [FindsBy(How = How.Id, Using = PURCHASE_ORDER_TAB)]
        private IWebElement _purchaseOrderTab;

        [FindsBy(How = How.Id, Using = PURCHASE_ORDER_NUMBER)]
        private IWebElement _purchaseOrderNumber;

        [FindsBy(How = How.Id, Using = PURCHASE_ORDER_DATE)]
        private IWebElement _purchaseOrderDate;
       
        [FindsBy(How = How.XPath, Using = PURCHASE_ORDER_SELECTED)]
        private IWebElement _purchaseOrderSelected;
        
        [FindsBy(How = How.Id, Using = SUBMIT)]
        private IWebElement _submit;

        //_______________________________ Méthodes ________________________________________________________

        public void FillField_CreatNewReceiptNotes(RNCreationParameters rnCP)
        {
            Random rnd = new Random();

            _site = WaitForElementIsVisibleNew(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, rnCP.SiteName);
            WaitForLoad();

            _supplier = WaitForElementIsVisibleNew(By.Id(SUPPLIER));
            _supplier.SetValue(ControlType.DropDownList, rnCP.SupplierName);

            _deliveryLocation = WaitForElementIsVisibleNew(By.Id(DELIVERY_LOCATION));
            _deliveryLocation.SetValue(ControlType.DropDownList, rnCP.DeliveryPlace);

            _deliveryOrderNumber = WaitForElementIsVisibleNew(By.Id(DELIVERY_ORDER_NUMBER));
            if (string.IsNullOrEmpty(rnCP.DeliveryNumber))
            {
                _deliveryOrderNumber.SetValue(ControlType.TextBox, rnd.Next(1000000, 9999999).ToString());
            }
            else
            {
                _deliveryOrderNumber.SetValue(ControlType.TextBox, rnCP.DeliveryNumber);
            }

            _activated = WaitForElementExists(By.Id(ACTIVATED));
            _activated.SetValue(ControlType.CheckBox, rnCP.IsActivated);

            _deliveryDate = WaitForElementIsVisibleNew(By.Id(DELIVERY_DATE));
            _deliveryDate.SetValue(ControlType.DateTime, rnCP.DeliveryDate);
            _deliveryDate.SendKeys(Keys.Tab);

            _createFrom = WaitForElementExists(By.Id(CREATE_FROM));
            _createFrom.SetValue(ControlType.CheckBox, rnCP.CreateFromPO);
            WaitForLoad();

            if (string.IsNullOrEmpty(rnCP.Commentary) == false)
            {
                var _commentaire = WaitForElementIsVisibleNew(By.Id(COMMENT));
                _commentaire.SetValue(ControlType.TextBox, rnCP.Commentary);
            }

            if (rnCP.CreateFromPO)
            {
                // Get purchase order
                _purchaseOrderTab = WaitForElementIsVisibleNew(By.Id(PURCHASE_ORDER_TAB));
                _purchaseOrderTab.Click();

                _purchaseOrderNumber = WaitForElementIsVisibleNew(By.Id(PURCHASE_ORDER_NUMBER));
                _purchaseOrderNumber.SetValue(ControlType.TextBox, rnCP.PONumber);

                _purchaseOrderDate = WaitForElementIsVisibleNew(By.Id(PURCHASE_ORDER_DATE));
                _purchaseOrderDate.SetValue(ControlType.DateTime, rnCP.PODate);
                WaitPageLoading();

                _purchaseOrderSelected = WaitForElementExists(By.XPath(string.Format(PURCHASE_ORDER_SELECTED, rnCP.PONumber)));
                _purchaseOrderSelected.SetValue(ControlType.CheckBox, true);
            }
            // attente de l'activation du bouton Submit en conséquence
            WaitForLoad();
        }

        public ReceiptNotesItem Submit()
        {
            _submit = WaitForElementIsVisibleNew(By.Id(SUBMIT));
            _submit.Click();
            WaitPageLoading();
            WaitForLoad();

            return new ReceiptNotesItem(_webDriver, _testContext);
        }

        public string GetDeliveryOrderNumber()
        {
            _deliveryOrderNumber = WaitForElementIsVisible(By.Id(DELIVERY_ORDER_NUMBER));
            return _deliveryOrderNumber.GetAttribute("value");
        }

        public string GetDeliveryDate()
        {
            _deliveryDate = WaitForElementIsVisible(By.Id(DELIVERY_DATE));
            return _deliveryDate.GetAttribute("value");
        }

        public bool isBackDatingValidateur()
        {
            string backDatingValidator = String.Format("//*/span[contains(@class,'text-danger') and contains(text(), '{0}')]", VALIDATOR_BACKDATING);
            return isElementVisible(By.XPath(backDatingValidator));
        }
    }

    public class RNCreationParameters
    {
        public DateTime DeliveryDate { get; set; } 
        public string SiteName { get; set; }
        public string SupplierName { get; set; }
        public string DeliveryPlace { get; set; }
        public string DeliveryNumber { get; set; }

        public bool IsActivated { get; set; }
        public bool CreateFromPO { get; set; }
        public string PONumber { get; set; }
        public DateTime? PODate { get; set; }
        public string Commentary { get; set; }

        public RNCreationParameters(DateTime deliveryDate, string siteName, string supplierName, string deliveryPlace)
        {
            IsActivated = true;
            CreateFromPO = false;

            DeliveryDate = deliveryDate;
            SiteName = siteName;
            SupplierName = supplierName;
            DeliveryPlace = deliveryPlace;
        }
    }
}
