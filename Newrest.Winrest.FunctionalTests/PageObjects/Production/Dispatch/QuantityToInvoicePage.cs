using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.Dispatch
{
    public class QuantityToInvoicePage : PageBase
    {
        public QuantityToInvoicePage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }


        //__________________________________Constantes_____________________________________

        private const string FIRST_ITEM = "//*[@id=\"dispatchTable\"]/tbody/tr[2]";
        private const string WEEK_QTY_INPUT = "/html/body/div[2]/div/div[2]/div[3]/div/div/div/table/tbody/tr[2]/td[{0}]/span[2]/input";
        private const string BTNVALIDATE = "0_ChkIsValidated";
        private const string QTY_INVO_SUNDAY_VALIDATION = "//*[@id=\"thSun\"]/span[2]/span[3]";
        private const string DAY_COLOR_VALIDATION = "/html/body/div[2]/div/div[2]/div[3]/div/div/div/table/tbody/tr[1]/th[{0}]/span[2]/span[1]";
        private const string ERROR_VALIDATION = "//*[@id=\"dataAlertModal\"]/div/div/div[2]/div/p";
        //__________________________________Variables______________________________________

        [FindsBy(How = How.XPath, Using = FIRST_ITEM)]
        private IWebElement _firstItem;

        [FindsBy(How = How.Id, Using = BTNVALIDATE)]
        private IWebElement _btnValidate;

        [FindsBy(How = How.Id, Using = ERROR_VALIDATION)]
        private IWebElement _errorValidate;

        [FindsBy(How = How.XPath, Using = QTY_INVO_SUNDAY_VALIDATION)]
        private IWebElement _qtyInvoSundayValidation;

        //___________________________________Pages_________________________________________

        public void UpdateQuantities(string newQty)
        {
            for (int i = 0; i < 12; i++)
            {
                try
                {
                    Actions action = new Actions(_webDriver);
                    action.MoveToElement(_firstItem).Perform();
                    var xpath = string.Format(WEEK_QTY_INPUT, i + 6);
                    var element = _webDriver.FindElement(By.XPath(xpath));
                    element.SetValue(ControlType.TextBox, newQty);
                    Thread.Sleep(1000);
                    WaitForLoad();

                }
                catch (Exception e)
                {
                    if (e.Message.Contains("Access Denied"))
                    {
                        MessageBox.Show("Sorry, Access Denied", "The input is not accessible");
                    }
                }
            }
        }

        public bool IsQuantitiesUpdated(string qty)
        {
            bool isUpdated = true;

            for (int i = 0; i < 7; i++)
            {
                try
                {
                    Actions action = new Actions(_webDriver);
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

        public void ValidateTheFirst()
        {
            Actions action = new Actions(_webDriver);
            action.MoveToElement(_firstItem).Perform();
            _btnValidate = WaitForElementIsVisible(By.Id(BTNVALIDATE));
            _btnValidate.Click();
            WaitForLoad();
        }

        public bool IsUpdateQty()
        {
            List<bool> isUp = new List<bool>();
            bool isOk = false;
            string newQty = "10";

            for (int i = 0; i < 12; i++)
            {
                try
                {
                    Actions action = new Actions(_webDriver);
                    action.MoveToElement(_firstItem).Perform();
                    var xpath = string.Format(WEEK_QTY_INPUT, i + 6);
                    var element = _webDriver.FindElement(By.XPath(xpath));
                    element.SetValue(ControlType.TextBox, newQty);
                    Thread.Sleep(1000);
                    WaitForLoad();
                    isUp.Add(true);
                }
                catch
                {
                    isUp.Add(false);

                }
            }

            if (isUp.Contains(true))
            {
                isOk = true;
            }
            else
            {
                isOk = false;
            }
            return isOk;
        }

        public bool IsSundayValidatedByColorDay()
        {
            bool isOk = false;

            try
            {
                _qtyInvoSundayValidation = _webDriver.FindElement(By.XPath(String.Format(QTY_INVO_SUNDAY_VALIDATION)));
                if (_qtyInvoSundayValidation.GetAttribute("style").Equals("border-bottom-color: green;"))
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
                    return  false;
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
    }
}
