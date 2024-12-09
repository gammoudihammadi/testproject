using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
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
   public  class ConvertersCreateModalPage : PageBase
    {
        private const string CLOSE_X = "//*[@id=\"modal-1\"]/div/div/form/div[1]/button/span";
        private const string CUSTOMER = "CustomerId";
        private const string CONVERTER_TYPE = "ConverterType";
        private const string FILE_EXTENSION = "FileExtension";
        private const string BNT_CREATE = "last";
        private const string VALIDATION_FILE_EXTENSION = "//*[@id='FileExtension']/..//span[contains(@class, 'text-danger col-md-12 field-validation-error')]";
        private const string VALIDATION_CONVERTER_TYPE = "//*[@id='ConverterType']/..//span[contains(@class, 'text-danger col-md-12 field-validation-error')]";
        private const string CUSTOMER_LABEL_XPATH = "//*[@id=\"SelectedCustomer-Group\"]/label";
        private const string CUSTOMER_INPUT_XPATH = "//*[@id=\"SelectedCustomer-Group\"]/div";
        private const string FILE_EXTENSION_LABEL_XPATH = "//*[@id=\"modal-1\"]/div/div/form/div[2]/div[2]/label";
        private const string FILE_EXTENSION_INPUT_XPATH = "//*[@id=\"modal-1\"]/div/div/form/div[2]/div[2]/div";
        private const string CONVERTER_TYPE_LABEL_XPATH = "//*[@id=\"modal-1\"]/div/div/form/div[2]/div[3]/label";
        private const string CONVERTER_TYPE_INPUT_XPATH = "//*[@id=\"modal-1\"]/div/div/form/div[2]/div[3]/div";

        public ConvertersCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
            
        }

        [FindsBy(How = How.Id, Using = CUSTOMER)]
        private IWebElement _customer;

        [FindsBy(How = How.Id, Using = CONVERTER_TYPE)]
        private IWebElement _converterType;

        [FindsBy(How = How.Id, Using = FILE_EXTENSION)]
        private IWebElement _fileExtension;

        [FindsBy(How = How.XPath, Using = BNT_CREATE)]
        private IWebElement _btn_create;

        public ConverterDetailsPage FillFiled_CreateConverters(string customer, string converterType, string fileExtension = null)
        {
            _customer = WaitForElementIsVisible(By.Id(CUSTOMER));
            _customer.SetValue(ControlType.DropDownList, customer);
            WaitForLoad();

            _converterType = WaitForElementIsVisible(By.Id(CONVERTER_TYPE));
            _converterType.SetValue(ControlType.DropDownList, converterType);
            WaitForLoad();

            if (fileExtension != null)
            {
                _fileExtension = WaitForElementIsVisible(By.Id(FILE_EXTENSION));
                _fileExtension.SetValue(ControlType.TextBox, fileExtension);
                WaitForLoad();
            }

            _btn_create = WaitForElementToBeClickable(By.Id(BNT_CREATE));
            _btn_create.Click();
            WaitForLoad();
            return new ConverterDetailsPage(_webDriver, _testContext);

        }

        public bool VerifyThreeLabelsAndInputsInSameRow()
        {
            var pairs = new List<(string labelXPath, string inputXPath)>
            {
              (CUSTOMER_LABEL_XPATH, CUSTOMER_INPUT_XPATH),
              (FILE_EXTENSION_LABEL_XPATH, FILE_EXTENSION_INPUT_XPATH),
              (CONVERTER_TYPE_LABEL_XPATH, CONVERTER_TYPE_INPUT_XPATH)
             };

            foreach (var (labelXPath, inputXPath) in pairs)
            {
                try
                {
                    IWebElement labelElement = WaitForElementIsVisible(By.XPath(labelXPath));
                    IWebElement inputElement = WaitForElementIsVisible(By.XPath(inputXPath));

                    IWebElement labelParent = labelElement.FindElement(By.XPath(".."));
                    IWebElement inputParent = inputElement.FindElement(By.XPath(".."));

                    if (labelParent.Equals(inputParent))
                    {
                        var labelLocation = labelElement.Location;
                        var inputLocation = inputElement.Location;

                        if (labelLocation.X >= inputLocation.X)
                        {
                            return false;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
                catch (NoSuchElementException)
                {
                    return false;
                }
            }
            return true;
        }


        public bool VerifyCreateValidators()
        {
            if (isElementVisible(By.XPath(VALIDATION_FILE_EXTENSION))
                && isElementVisible(By.XPath(VALIDATION_CONVERTER_TYPE)))
            {
                return true;
            }
            return false;
        }

    }
}
