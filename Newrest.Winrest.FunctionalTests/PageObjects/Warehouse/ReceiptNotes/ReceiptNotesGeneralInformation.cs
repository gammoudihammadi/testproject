using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Purchasing;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.Inventory;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Warehouse.ReceiptNotes
{
    public class ReceiptNotesGeneralInformation : PageBase
    {
        public ReceiptNotesGeneralInformation(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //_____________________________ Constantes ______________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        // Onglets
        private const string ITEMS_TAB = "hrefTabContentItems";
        private const string PURCHASE_ORDER_PAGE = "//*[@id=\"form-create-receipt-note\"]/div/div[2]/div/div/div/span/a";
        private const string PURCHASE_ORDER_PAGE_2 = "//*[@id=\"form-create-receipt-note\"]/div/div[2]/div/div/div/span/a[2]";

        // Informations
        private const string RECEIPT_NOTE_NUMBER = "tb-new-receipt-note-number";
        private const string CLAIM_NUMBER = "//*[@id=\"form-create-receipt-note\"]/div/div[2]/div/div/div/span/a";
        private const string RECEIPT_NOTE_DELIVERY_ORDER_NUMBER = "ReceiptNote_DeliveryOrderNumber";
        private const string RECEIPT_NOTE_COMMENT = "ReceiptNote_Comment";
        private const string RECEIPT_NOTE_DELIVERYDATE = "datapicker-new-receipt-note-delivery";

        private const string STATUS = "ReceiptNote_Status";       
        private const string IS_ACTIVE = "ReceiptNote_IsActive";

        //_____________________________ Variables _______________________________________

        // General
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        // Onglets
        [FindsBy(How = How.Id, Using = ITEMS_TAB)]
        private IWebElement _itemsTab;

        [FindsBy(How = How.XPath, Using = PURCHASE_ORDER_PAGE)]
        private IWebElement _purchaseOrderPage;

        // Informations
        [FindsBy(How = How.Id, Using = RECEIPT_NOTE_NUMBER)]
        private IWebElement _receiptNoteNumber;

        [FindsBy(How = How.Id, Using = RECEIPT_NOTE_DELIVERY_ORDER_NUMBER)]
        private IWebElement _receiptNoteDeliveryOrderNumber;
        

        [FindsBy(How = How.XPath, Using = CLAIM_NUMBER)]
        private IWebElement _claimNumber;

        [FindsBy(How = How.Id, Using = STATUS)]
        private IWebElement _status;
       
        [FindsBy(How = How.Id, Using = IS_ACTIVE)]
        private IWebElement _isActive;

        //___________________________ Methodes ________________________________________

        // General
        public ReceiptNotesPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new ReceiptNotesPage(_webDriver, _testContext);
        }

        // Onglets
        public ReceiptNotesItem ClickOnItemsTab()
        {
            _itemsTab = WaitForElementIsVisible(By.Id(ITEMS_TAB));
            _itemsTab.Click();
            WaitForLoad();

            return new ReceiptNotesItem(_webDriver, _testContext);
        }

        public PurchaseOrderItem ClickOnPurchaseOrderLink()
        {
            _purchaseOrderPage = WaitForElementIsVisible(By.XPath(PURCHASE_ORDER_PAGE));
            _purchaseOrderPage.Click();
            WaitForLoad();

            return new PurchaseOrderItem(_webDriver, _testContext);
        }

        // Informations
      
        public string GetReceiptNoteNumber()
        {
            _receiptNoteNumber = WaitForElementExists(By.Id(RECEIPT_NOTE_NUMBER));
            return _receiptNoteNumber.GetAttribute("value");
        }

        public string GetReceiptNoteDeliveryOrderNumber()
        {
            _receiptNoteDeliveryOrderNumber = WaitForElementExists(By.Id(RECEIPT_NOTE_DELIVERY_ORDER_NUMBER));
            return _receiptNoteDeliveryOrderNumber.GetAttribute("value");
        }

        public string GetClaimNumber()
        {
            _claimNumber = WaitForElementExists(By.XPath(CLAIM_NUMBER));
            return _claimNumber.Text;
        }

        public void SetStatus(string status)
        {
            _status = WaitForElementIsVisible(By.Id(STATUS));
            _status.SetValue(ControlType.DropDownList, status);

            WaitForLoad();
        }

        public bool GetGeneralInfoActivateState()
        {
            _isActive = WaitForElementExists(By.Id(IS_ACTIVE));
            return _isActive.Selected;
        }

        public void SetGeneralInfoActivateState(bool value)
        {
            _isActive = WaitForElementExists(By.Id(IS_ACTIVE));
            _isActive.SetValue(ControlType.CheckBox, value);
        }

        public void Fill(DateTime date, string deliveryOrderNumber, string comment)
        {
            var rnDelivery = WaitForElementIsVisible(By.Id("ReceiptNote_DeliveryOrderNumber"));
            rnDelivery.Clear();
            rnDelivery.SendKeys(deliveryOrderNumber);

            var rnDate = WaitForElementIsVisible(By.Id("datapicker-new-receipt-note-delivery"));
            rnDate.Clear();
            rnDate.SendKeys(date.ToString("dd/MM/yyyy"));

            var rnComment = WaitForElementIsVisible(By.Id("ReceiptNote_Comment"));
            rnComment.Clear();
            rnComment.SendKeys(comment);

            //data binding
            Thread.Sleep(2000);
            WaitForLoad();
        }

        internal string GetDeliveryOrderNumber()
        {
            var rnDelivery = WaitForElementIsVisible(By.Id("ReceiptNote_DeliveryOrderNumber"));
            return rnDelivery.GetAttribute("value");
        }

        public string GetDate()
        {
            var rnDate = WaitForElementIsVisible(By.Id("datapicker-new-receipt-note-delivery"));
            return rnDate.GetAttribute("value");
        }

        public string GetComment()
        {
            var rnComment = WaitForElementIsVisible(By.Id("ReceiptNote_Comment"));
            return rnComment.Text;
        }

        public string GetPurchaseOrderNumber()
        {
            _purchaseOrderPage = WaitForElementIsVisible(By.XPath(PURCHASE_ORDER_PAGE));
            return _purchaseOrderPage.Text;
        }
        public string GetPurchaseOrderNumbers()
        {
            var _purchaseOrders = _webDriver.FindElements(By.XPath(PURCHASE_ORDER_PAGE));
            string liens = "";
            foreach (var p in _purchaseOrders)
            {
                if (liens == "")
                {
                    liens += p.Text;
                } else
                {
                    liens += " " + p.Text;
                }
            }
            return liens;
        }

        public string GetStatus()
        {
            var statusSelect = WaitForElementIsVisible(By.XPath("//*/select[@id='ReceiptNote_Status']/option[@selected]"));
            return statusSelect.Text;
        }

        public string GetDateValidated()
        {
            if (IsDev())
            {
                var date = WaitForElementIsVisible(By.Id("datapicker-new-receipt-note-delivery"));
                return date.GetAttribute("value");
            }
            else
            {
                var date = WaitForElementIsVisible(By.XPath("//*/label[contains(text(),'Delivery date : ')]/../div/input"));
                return date.GetAttribute("value");
            }     
        }

        public void SetReceiptNoteDeliveryOrderNumber(string value)
        {
            var rnDeliveryOrderNumber = WaitForElementIsVisible(By.Id(RECEIPT_NOTE_DELIVERY_ORDER_NUMBER));
            //new Actions(_webDriver).MoveToElement(rnDeliveryOrderNumber).Perform();
            rnDeliveryOrderNumber.SetValue(ControlType.TextBox, value);
            WaitForLoad();
        }

        public void SetComment(string comment)
        {
            var rnComment = WaitForElementExists(By.Id(RECEIPT_NOTE_COMMENT));
            new Actions(_webDriver).MoveToElement(rnComment).Perform();
            rnComment = WaitForElementIsVisible(By.Id(RECEIPT_NOTE_COMMENT));
            // TextArea
            rnComment.Clear();
            rnComment.SendKeys(comment);
            WaitForLoad();
        }

        public void SetDeliveryDate(DateTime dateTime)
        {
            var rnDeliveryDate = WaitForElementIsVisible(By.Id(RECEIPT_NOTE_DELIVERYDATE));
            //new Actions(_webDriver).MoveToElement(rnDeliveryDate).Perform();
            rnDeliveryDate.SetValue(ControlType.DateTime, dateTime);
            WaitForLoad();
        }

        public void PageUp()
        {
            var rnTitle = WaitForElementExists(By.XPath("/html/body/div[3]/div/div[1]/h1"));
            new Actions(_webDriver).MoveToElement(rnTitle).Perform();
            WaitForLoad();
            Thread.Sleep(2000);
            WaitForLoad();
        }
    }
}
