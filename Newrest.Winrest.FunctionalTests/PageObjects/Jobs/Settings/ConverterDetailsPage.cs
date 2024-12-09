using DocumentFormat.OpenXml.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Catalogs;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings
{
    public  class ConverterDetailsPage: PageBase
    {
        private const string DETAILS = "tabGeneralInformation";
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string CLOSE_X = "//*[@id=\"modal-1\"]/div/div/form/div[1]/button/span";
        private const string CUSTOMER = "CustomerId";
        private const string CONVERTER_TYPE = "ConverterType";
        private const string BTN_SAVE = "last"; // Assumed to be the same button for saving
        private const string FIRST_ROW = "//*[@id=\"tableListMenu\"]/tbody/tr";
        private const string TAB_CONV_SETTINGS = "tabConverterSettings";
        private const string NEW_CONVERTER_BTN = "/html/body/div[3]/div/div/div[2]/div/div/div/div/div/a[2]";
        private const string SELECT_MISSING_KEY = "drop-down-missing-keys";
        private const string VALUE = "Value";
        private const string CREATE = "last";
        private const string CREATED_VALUE = "/html/body/div[3]/div/div/div[2]/div/div/table/tbody/tr/td[3]";
        private const string KEY_INCLUDED = "/html/body/div[4]/div/div/div/div/form/div[2]/div[1]/div/div[2]/select/option[13]"; 

        public ConverterDetailsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
                
        }
        [FindsBy(How = How.Id, Using = DETAILS)]
        private IWebElement _details;

        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.Id, Using = CUSTOMER)]
        private IWebElement _customer;

        [FindsBy(How = How.Id, Using = CONVERTER_TYPE)]
        private IWebElement _converterType;

        [FindsBy(How = How.XPath, Using = BTN_SAVE)]
        private IWebElement _btn_save;

        [FindsBy(How = How.Id, Using = VALUE)]
        private IWebElement _value;

        [FindsBy(How = How.Id, Using = CREATE)]
        private IWebElement _create;
        

        public void ClickOnBackToList()
        {
            var backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            backToList.Click();

        }
        public SettingsTabConvertersPage BackToList()
        {
            _backToList.Click();
            WaitForLoad();
            return new SettingsTabConvertersPage(_webDriver, _testContext);
        }

        public void FillFields_EditConverters(string customer, string converterType)
        {

            _customer = WaitForElementIsVisible(By.Id(CUSTOMER));
            _customer.SetValue(ControlType.DropDownList, customer);

            WaitPageLoading();

            _converterType = WaitForElementIsVisible(By.Id(CONVERTER_TYPE));
            _converterType.SetValue(ControlType.DropDownList, converterType);


            WaitForLoad();
        }
        public SettingsTabConvertersPage SelectFirstRow()
        {
            WaitForLoad();
            var firstRow = WaitForElementIsVisible(By.XPath(FIRST_ROW));
            firstRow.Click();
            WaitPageLoading();
            return new SettingsTabConvertersPage(_webDriver, _testContext);
        }

        public void GoToSettingsConverter()
        {
            var _converterSettings = WaitForElementExists(By.Id(TAB_CONV_SETTINGS));
            _converterSettings.Click();
            WaitForLoad();
        }
        public ConverterDetailsPage ClickNewSettingCreatePage()
        {
            WaitForLoad();
            var _newConverterBtn = WaitForElementIsVisible(By.XPath(NEW_CONVERTER_BTN));
            _newConverterBtn.Click();
            WaitPageLoading();
            return new ConverterDetailsPage(_webDriver, _testContext);
        }

        public ConverterDetailsPage FillField_CreatNewSettingConverter(string value , string IncludedPacket)
        {
         
            var s = WaitForElementIsVisible(By.Id(SELECT_MISSING_KEY));
            s.SetValue(ControlType.TextBox, IncludedPacket);
            WaitForLoad();

            if (isElementExists(By.Id(SELECT_MISSING_KEY)))
            {
                var selectKey = WaitForElementIsVisible(By.XPath(KEY_INCLUDED));
                selectKey.Click();
                WaitForLoad();
            }
            s.Click();

            if (value != null)
            {
                _value = WaitForElementIsVisible(By.Id(VALUE));
                _value.SetValue(ControlType.TextBox, value);
                WaitForLoad();
            }

            _create = WaitForElementIsVisible(By.Id(CREATE));
            _create.Click();

            return new ConverterDetailsPage(_webDriver, _testContext);
        }


        public bool IscreatedOKSettingsConverter(string value)
        {
             
                WaitPageLoading();
                var createdValue = _webDriver.FindElement(By.XPath(CREATED_VALUE));
                WaitPageLoading();
                string valueKey = createdValue.Text;
                WaitPageLoading();


            if (valueKey == value)
            {
                return true;
            }
            else return false; 
        }
    }
}
