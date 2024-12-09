using GemBox.Email.Calendar;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Interim.InterimReceptions
{
    public class InterimReceptionsGeneralInformation : PageBase
    {
        private const string INTERIM_NUMBER = "tb-new-interim-reception-number";
        private const string COMMENT = "InterimReception_Comment";

        private const string SITE = "InterimReception_SitePlace_Site_Name";
        private const string SUPPLIER = "InterimReception_Supplier_Name";
        private const string DELIVERY_LOCATION = "InterimReception_SitePlace_Title";
        private const string DELIVERY_ORDER_NUMBER = "InterimReception_DeliveryOrderNumber";
        private const string DELIVERY_DATE = "//*[@id=\"form-create-interim-reception\"]/div/div[5]/div/div/div/input";
        private const string VALIDATED_BY = "InterimReception_UserValidator_FullName";

        //__________________________________ Variables _________________________________________________


        [FindsBy(How = How.Id, Using = INTERIM_NUMBER)]
        private IWebElement _interimNumber;
        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;
        [FindsBy(How = How.Id, Using = SUPPLIER)]
        private IWebElement _supplier;
        [FindsBy(How = How.Id, Using = DELIVERY_LOCATION)]
        private IWebElement _delivery_location;
        [FindsBy(How = How.Id, Using = DELIVERY_ORDER_NUMBER)]
        private IWebElement _delivery_order_number;
        [FindsBy(How = How.XPath, Using = DELIVERY_DATE)]
        private IWebElement _delivery_date;
         [FindsBy(How = How.Id, Using = VALIDATED_BY)]
        private IWebElement _validated_by;


        public InterimReceptionsGeneralInformation(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Methodes_______________________________________

        public string GetInterimId()
        {
            _interimNumber = WaitForElementIsVisible(By.Id(INTERIM_NUMBER));
            _interimNumber.Click();
            return _interimNumber.GetAttribute("value");
        }
        public string GetComment()
        {
            var comment = WaitForElementIsVisible(By.Id(COMMENT)); 
            return comment.Text;
        }
        public string GetSite()
        {
            _site = WaitForElementIsVisible(By.Id(SITE));
            //_site.Click();
            return _site.GetAttribute("value");
        }
        public string GetSupplier()
        {
            _supplier = WaitForElementIsVisible(By.Id(SUPPLIER));
            //_supplier.Click();
            return _supplier.GetAttribute("value");
        }
        public string GetDelivery_location()
        {
            _delivery_location = WaitForElementIsVisible(By.Id(DELIVERY_LOCATION));
           // _delivery_location.Click();
            return _delivery_location.GetAttribute("value");
        }
        public string GetDelivery_order_Number()
        {
            _delivery_order_number = WaitForElementIsVisible(By.Id(DELIVERY_ORDER_NUMBER));
            //_delivery_order_number.Click();
            return _delivery_order_number.GetAttribute("value");
        }
        public string GetDelivery_Date()
        {
            _delivery_date = WaitForElementIsVisible(By.XPath(DELIVERY_DATE));
            //_delivery_date.Click();
            return _delivery_date.GetAttribute("value");
        }
        public string GetValidated_By()
        {
            _validated_by = WaitForElementIsVisible(By.Id(VALIDATED_BY));
            //_validated_by.Click();
            return _validated_by.GetAttribute("value");
        }
    }
}
