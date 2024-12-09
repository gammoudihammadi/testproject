using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class CreatePurchaseOrderModalPage : PageBase
    {

        public CreatePurchaseOrderModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _______________________________________ Constantes _______________________________________________

        private const string PURCHASE_ORDER_NUMBER = "tb-new-purchase-order-number";
        private const string SITE = "SelectedSiteId";
        private const string SUPPLIER = "drop-down-suppliers";
        private const string DELIVERY_LOCATION = "SelectedSitePlaceId";
        private const string DELIVERY_DATE = "datapicker-new-purchase-order-delivery";
        private const string ACTIVATED = "PurchaseOrder_IsActive";
        private const string CHECKBOX_COPY_FROM = "checkBoxCopyFrom";
        private const string FROM_PO_NUMBER = "form-copy-from-tbSearchName";
        private const string FIRST_ITEM_SELECTED = "//*[@id=\"table-copy-from-po\"]/tbody/tr[2]/td[6]/div";
        private const string SUBMIT = "btn-submit-form-create-purchase-order";
        private const string BACK_TO_LIST = "//html/body/div[2]/a/span[1]";


        // _______________________________________ Variables ________________________________________________

        [FindsBy(How = How.Id, Using = PURCHASE_ORDER_NUMBER)]
        private IWebElement _purchaseOrderNumber;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;

        [FindsBy(How = How.Id, Using = SUPPLIER)]
        private IWebElement _supplier;

        [FindsBy(How = How.Id, Using = DELIVERY_LOCATION)]
        private IWebElement _deliveryLocation;

        [FindsBy(How = How.Id, Using = DELIVERY_DATE)]
        private IWebElement _deliveryDate;

        [FindsBy(How = How.Id, Using = ACTIVATED)]
        private IWebElement _activated;

        [FindsBy(How = How.Id, Using = CHECKBOX_COPY_FROM)]
        private IWebElement _checkBoxCopyFrom;

        [FindsBy(How = How.Id, Using = FROM_PO_NUMBER)]
        private IWebElement _fromPONumber;

        [FindsBy(How = How.XPath, Using = FIRST_ITEM_SELECTED)]
        private IWebElement _itemSelected;

        [FindsBy(How = How.Id, Using = SUBMIT)]
        private IWebElement _submit;      

        // _______________________________________ Méthodes __________________________________________________

       public void FillPrincipalField_CreationNewPurchaseOrder(string site, string supplier, string location, DateTime deliveryDate, bool activate, bool copyFromOtherPO = false, string poNumber = null)
        {
            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site);
            WaitForLoad();

            _supplier = WaitForElementIsVisible(By.Id(SUPPLIER));
            _supplier.SetValue(ControlType.DropDownList, supplier);

            _deliveryLocation = WaitForElementIsVisible(By.Id(DELIVERY_LOCATION));
            _deliveryLocation.SetValue(ControlType.DropDownList, location);

            _deliveryDate = WaitForElementIsVisible(By.Id(DELIVERY_DATE));
            _deliveryDate.SetValue(ControlType.DateTime, deliveryDate);
            _deliveryDate.SendKeys(Keys.Tab);
            _activated = WaitForElementExists(By.Id("IsActive"));
            _activated.SetValue(ControlType.CheckBox, activate);            

            if(copyFromOtherPO)
            {
                // On coche la checkBox
                _checkBoxCopyFrom = WaitForElementExists(By.Id(CHECKBOX_COPY_FROM));
                _checkBoxCopyFrom.SetValue(ControlType.CheckBox, true);
                WaitForLoad();

                _fromPONumber = WaitForElementIsVisible(By.Id(FROM_PO_NUMBER));
                _fromPONumber.SetValue(ControlType.TextBox, poNumber);
                WaitPageLoading();
                Thread.Sleep(1000);

                _itemSelected = WaitForElementToBeClickable(By.XPath(FIRST_ITEM_SELECTED));
                _itemSelected.SetValue(ControlType.CheckBox, true);
            }
        }

        public PurchaseOrderItem Submit()
        {
            _submit = WaitForElementToBeClickable(By.Id(SUBMIT));
            _submit.Click();
            WaitForLoad();

            return new PurchaseOrderItem(_webDriver, _testContext);
        }

        /*        public string GetPurchaseOrderNumber()
            {
                _purchaseOrderNumber = WaitForElementIsVisible(By.Id(PURCHASE_ORDER_NUMBER));
                return _purchaseOrderNumber.GetAttribute("value");
            }*/
        public void ClickOnBackToList()
        {
            var backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            backToList.Click();
            WaitForLoad();
        }

    }
}
