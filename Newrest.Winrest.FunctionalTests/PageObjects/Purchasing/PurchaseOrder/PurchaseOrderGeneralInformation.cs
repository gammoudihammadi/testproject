using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Globalization;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class PurchaseOrderGeneralInformation : PageBase
    {

        public PurchaseOrderGeneralInformation(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ____________________________________ Constantes _____________________________________________

        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string ITEMS_TAB = "hrefTabContentItems";

        private const string DELIVERY_DATE = "datapicker-new-purchase-order-delivery";
        private const string COMMENT = "PurchaseOrder_Comment";
        private const string STATUS = "PurchaseOrder_Status";
        private const string STATUS_SELECTED = "//*[@id=\"PurchaseOrder_Status\"]/option[@selected = 'selected']";
        private const string PURCHASE_ORDER_NUMBER = "//*[@id='div-body']/div/div[1]/h1";
        private const string USERVALIDATOR_IMPUT = "UserValidatorFullName";
        private const string APPROVED_BY_INPUT = "#collapseApproval input";
        private const string SEND_BY_EMAIL = "btn-send-by-email-purchase-order";
        // ____________________________________ Variables ______________________________________________

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList; 

         [FindsBy(How = How.Id, Using = ITEMS_TAB)]
        private IWebElement _itemsTab;

        [FindsBy(How = How.Id, Using = DELIVERY_DATE)]
        private IWebElement _deliveryDate;

        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.Id, Using = STATUS)]
        private IWebElement _status;

        [FindsBy(How = How.XPath, Using = STATUS_SELECTED)]
        private IWebElement _statusSelected;

        [FindsBy(How = How.XPath, Using = PURCHASE_ORDER_NUMBER)]
        private IWebElement _purchaseOrderNumber;

        [FindsBy(How = How.CssSelector, Using = APPROVED_BY_INPUT)]
        private IWebElement _approvedByInput;

        [FindsBy(How = How.Id, Using = USERVALIDATOR_IMPUT)]
        private IWebElement _userValidatorInput;

        // _____________________________________ Méthodes ______________________________________________

        public PurchaseOrdersPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new PurchaseOrdersPage(_webDriver, _testContext);
        }

        public string GetDeliveryDateValue()
        {
            _deliveryDate = WaitForElementIsVisible(By.Id(DELIVERY_DATE));

            var dateFormat = _deliveryDate.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            return DateTime.Parse(_deliveryDate.GetAttribute("value"), ci).ToShortDateString();
        }

        public void DeliveryDateUpdate(DateTime date)
        {
            _deliveryDate = WaitForElementIsVisible(By.Id(DELIVERY_DATE));
            _deliveryDate.SetValue(ControlType.DateTime, date);
            _deliveryDate.SendKeys(Keys.Tab);
            WaitForLoad();
        }

        public string GetComment()
        {
            _comment = WaitForElementIsVisible(By.Id("Comment"));
            return _comment.GetAttribute("value");
        }

        public void SetComment(string comment)
        {
            _comment = WaitForElementIsVisible(By.Id("Comment"));
            _comment.SetValue(ControlType.TextBox, comment);
        }

        public string GetStatus()
        {
            _statusSelected = WaitForElementExists(By.XPath("//*[@id=\"PurchaseOrderStatus\"]/option[@selected = 'selected']"));
            return _statusSelected.Text;
        }

        public void SetStatus(string status)
        {
            _status = WaitForElementIsVisible(By.Id("PurchaseOrderStatus"));
            _status.SetValue(ControlType.DropDownList, status);

            Thread.Sleep(2000);
        }

        public PurchaseOrderItem ClickOnItemsTab()
        {
            _itemsTab = WaitForElementIsVisible(By.Id(ITEMS_TAB));
            _itemsTab.Click();
            WaitForLoad();

            return new PurchaseOrderItem(_webDriver, _testContext);
        }

        public string getPurchaseOrderNumber()
        {
            _purchaseOrderNumber = WaitForElementIsVisible(By.XPath(PURCHASE_ORDER_NUMBER));
            return _purchaseOrderNumber.Text.Substring("PURCHASE ORDER NO°".Length);
        }

        public string getSupplyOrderNumber()
        {
            IWebElement _supplyOrderNumber = null;
            _supplyOrderNumber = WaitForElementExists(By.XPath("//*[@id='form-create-purchase-order']/div/div[9]/div/div/div/span/a"));
            return _supplyOrderNumber.Text;
        }
        public string getSupplyOrderNumberValidated()
        {
            IWebElement _supplyOrderNumber = null;
            _supplyOrderNumber = WaitForElementExists(By.XPath("//*[@id='form-create-purchase-order']/div/div[7]/div/div/div/span/a"));
            return _supplyOrderNumber.Text;
        }

        public void SetActivated(bool value)
        {
            var isActive = WaitForElementExists(By.Id("IsActive"));
            isActive.SetValue(PageBase.ControlType.CheckBox, value);
            Thread.Sleep(2000);
            WaitForLoad();
        }

        public bool GetActivated()
        {
            var isActive = WaitForElementExists(By.Id("IsActive"));
            return isActive.GetAttribute("checked") != null;
        }

        public string GetUserValidator()
        {
            _userValidatorInput = WaitForElementIsVisible(By.Id(USERVALIDATOR_IMPUT));
            return _userValidatorInput.GetAttribute("value");
        }

        public string GetApprovedBy()
        {
            _approvedByInput = WaitForElementIsVisible(By.CssSelector(APPROVED_BY_INPUT));
            return _approvedByInput.GetAttribute("value");
        }
        public bool GetGenenralInfoEmail(string ID , string Date)
        {
            var sendBtn = WaitForElementIsVisible(By.Id(SEND_BY_EMAIL));
            sendBtn.Click();
    
            var sub = WaitForElementIsVisible(By.Id("Subject"));
           var text =  sub.GetAttribute("value");
            return text.Contains(ID) && text.Contains(Date); 
           
        }
    }
}
