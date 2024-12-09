using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.Dispatch
{
    public class QuantityToProducePage : PageBase
    {

        public QuantityToProducePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_____________________________________

        private const string FIRST_ITEM = "//*[@id=\"dispatchTable\"]/tbody/tr[2]";
        private const string WEEK_QTY_INPUT = "/html/body/div[2]/div/div[2]/div[3]/div/div/div/table/tbody/tr[2]/td[{0}]/span[2]/input";
        private const string BTNVALIDATE = "0_ChkIsValidated";
        private const string QTY_PROD_SUNDAY_VALIDATION = "//*[@id=\"thSun\"]/span[2]/span[2]";
        private const string DAY_COLOR_VALIDATION = "/html/body/div[2]/div/div[2]/div[3]/div/div/div/table/tbody/tr[1]/th[{0}]/span[2]/span[1]";
        private const string ERROR_VALIDATION = "//*[@id=\"dataAlertModal\"]/div/div/div[2]/div/p";
        private const string SEARCH_FILTER = "//*[@id=\"SearchPattern\"]";
        private const string UPDATE_QTYS_DISPLAY = "/html/body/div[2]/div/div[2]/div[3]/div/div/div/table/tbody/tr[2]/td[{0}]/span[1]";
        private const string UPDATE_QTYS_EDIT = "/html/body/div[2]/div/div[2]/div[3]/div/div/div/table/tbody/tr[2]/td[{0}]/span[2]";

        //__________________________________Variables______________________________________

        [FindsBy(How = How.XPath, Using = SEARCH_FILTER)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.XPath, Using = FIRST_ITEM)]
        private IWebElement _firstItem;

        [FindsBy(How = How.Id, Using = BTNVALIDATE)]
        private IWebElement _btnValidate;

        [FindsBy(How = How.Id, Using = ERROR_VALIDATION)]
        private IWebElement _errorValidate;

        [FindsBy(How = How.XPath, Using = QTY_PROD_SUNDAY_VALIDATION)]
        private IWebElement _qtyProdSundayValidation;


        //___________________________________Pages_________________________________________

        public void UpdateQuantities(string newQty)
        {
            if (isElementVisible(By.XPath(String.Format(UPDATE_QTYS_DISPLAY, 6))))
            {
                // passer en mode édition
                var element = WaitForElementIsVisible(By.XPath(String.Format(UPDATE_QTYS_DISPLAY, 6)));
                element.Click();
            }
            for (int i = 0; i < 7; i++)
            {
                var clickElement = WaitForElementIsVisible(By.XPath(String.Format(UPDATE_QTYS_EDIT, i + 6)));
                clickElement.Click();
                var element = WaitForElementIsVisible(By.XPath(String.Format(WEEK_QTY_INPUT, i + 6)));
                element.Clear();
                element.SetValue(ControlType.TextBox, newQty);
                
                
                WaitPageLoading();
            }
        }

        public string GetQuantityToProduceForTomorrow()
        {
            var caseInput = WaitForElementIsVisible(By.XPath("//input[contains(@id,'Qty" + DateUtils.Now.AddDays(1).DayOfWeek + "')]"));
            return caseInput.GetAttribute("value");
        }

        public bool IsQuantitiesUpdated(string qty, int from=0, int to=6)
        {
            bool isUpdated = true;

            for (int i = from; i <= to; i++)
            {
                try
                {
                    Actions action = new Actions(_webDriver);
                    _firstItem = WaitForElementIsVisible(By.XPath(FIRST_ITEM));
                    action.MoveToElement(_firstItem).Perform();

                    var element = _webDriver.FindElement(By.XPath(String.Format(WEEK_QTY_INPUT, i + 6)));

                    if (!element.GetAttribute("value").Equals(qty))
                    {
                        isUpdated = false;
                        break;
                    }
                }
                catch
                {
                    isUpdated = false;
                }
            }

            return isUpdated;
        }

        public void ValidateFirstDispatch()
        {
            Actions action = new Actions(_webDriver);
            action.MoveToElement(_firstItem).Perform();

            _btnValidate = WaitForElementExists(By.Id(BTNVALIDATE));
            _btnValidate.Click();
            WaitForLoad();
        }

        public bool IsDispatchValidated()
        {
            _btnValidate = WaitForElementIsVisible(By.Id(BTNVALIDATE));
            if (_btnValidate.GetAttribute("data-isvalidated").Equals("true"))
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public bool CanUpdateQty()
        {
            bool isOk = false;

            for (int i = 0; i < 7; i++)
            {
                try
                {
                    Actions action = new Actions(_webDriver);
                    action.MoveToElement(_firstItem).Perform();

                    var element = _webDriver.FindElement(By.XPath(string.Format(WEEK_QTY_INPUT, i + 6)));
                    if (element.GetAttribute("disabled") == null)
                    {
                        isOk = true;
                        break;
                    }
                }
                catch
                {
                    isOk = false;
                }
            }

            return isOk;
        }

        public bool IsSundayValidatedByColorDay()
        {
            bool isOk = false;

            try
            {
                _qtyProdSundayValidation = _webDriver.FindElement(By.XPath(String.Format(QTY_PROD_SUNDAY_VALIDATION)));
                if (_qtyProdSundayValidation.GetAttribute("style").Equals("border-bottom-color: green;"))
                    isOk = true;
            }
            catch
            {
                isOk = false;
            }

            return isOk;
        }

        public bool IsValidatedByColorDay()
        {
            bool isOk = true;

            for (int i = 0; i < 7; i++)
            {
                try
                {
                    var element = _webDriver.FindElement(By.XPath(String.Format(DAY_COLOR_VALIDATION, i + 6)));

                    if (!element.GetAttribute("style").Equals("border-bottom-color: green;"))
                    {
                        isOk = false;
                        break;
                    }
                }
                catch
                {
                    isOk = false;
                }
            }

            return isOk;
        }

        public bool IsErrorValidation()
        {
          
                if (isElementVisible(By.XPath(ERROR_VALIDATION)))
                {
                    return true;
                }
                else
                {
                    return false;
                }
        
        }

        public void searchDispathByServiceName(string serviceName)
        {
            _searchFilter = WaitForElementExists(By.XPath(SEARCH_FILTER));
            _searchFilter.SetValue(ControlType.TextBox, serviceName);
            WaitForLoad();
        }
    }
}
