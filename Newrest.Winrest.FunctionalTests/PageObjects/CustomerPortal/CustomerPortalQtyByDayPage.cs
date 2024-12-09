using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OpenQA.Selenium;
using System.Globalization;
using System;
using System.Security.Policy;
using OpenQA.Selenium.Interactions;
using System.Threading;
using Newrest.Winrest.FunctionalTests.Utils;
using DocumentFormat.OpenXml.Bibliography;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using System.Collections.Generic;
using System.Linq;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.CustomerPortal
{
    public class CustomerPortalQtyByDayPage : CustomerPortalQtyByWeekPage
    {

        private const string NEXT_DAY = "//*[@id=\"selectDate\"]/a[2]/span";
        private const string PREVIOUS_DAY = "//*[@id=\"selectDate\"]/a[1]/span";
        private const string CURRENT_CALENDAR_DAY = "//*[@id=\"selectDate\"]/span[2]";
        private const string SERVICES = "//*[@id=\"tabContentItemContainer\"]/div[2]/div[2]/div/table/tbody/tr[*]/td[1]";
        private const string ERROR_MSG = "//*[@id=\"tabContentItemContainer\"]/div[2]/div[2]/div/p";
        private const string MOVE_FORWARD_BTN = "/html/body/div[3]/div/div[2]/div[1]/div[2]/div[3]/a[2]/span";
        private const string MOVE_BACKWARD_BTN = "/html/body/div[3]/div/div[2]/div[1]/div[2]/div[3]/a[1]/span";
        private const string QTY_INPUT = "//*[@id=\"New-60179-894687-532QtyMonday\"]";
        private const string QTY_INPUT_BY_DAY = "//*[@id=\"tabContentItemContainer\"]/div[2]/div[2]/div/table/tbody/tr[*][@data-step='Production']/td[1][contains(text(),'{0}')]/../td/div//*[contains(@id,\"Qty{1}\")]";
        private const string COMMENT_BTN = "//*[@id=\"tabContentItemContainer\"]/div[2]/div[2]/div/table/tbody/tr[4]/td[2]/div/a";
        private const string COMMENT_AREA = "textarea-comment";
        private const string SUBMIT_BTN = "//*[@id=\"form-save-comment\"]/div[2]/button[1]";
        private const string VALIDATE = "btn-validate-CP-1";
        private const string QTY = "/html/body/div[3]/div/div[2]/div[2]/div[2]/div/table/tbody/tr[4]/td[2]/div/p";
        private const string COMMENT_VALUE = "//*[@id=\"form-save-comment\"]/div[1]/p[2]";
        private const string QTY_BY_DATE = "/html/body/div[3]/div/div[2]/div[2]/div[2]/div/table/tbody/tr[4]/td[2]/div/p/input";
        private const string SERVICE_NUMBER = "/html/body/div[3]/div/div[2]/div[2]/div[2]/div/table/tbody/tr[4]/td[2]/div/p/input";
        private const string DAY_NAME = "//*[@id=\"th1\"]/span[1]";

        public CustomerPortalQtyByDayPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public void ClickOnNextDay()
        {
           var nextDay = WaitForElementIsVisible(By.XPath(NEXT_DAY));
            nextDay.Click();
            WaitForLoad();
        }

        public string GetDayName()
        {
            var nextDay = WaitForElementIsVisible(By.XPath(DAY_NAME));
            return nextDay.Text;
        }
        public void ClickOnPreviousDay()
        {
            var previousDay = WaitForElementIsVisible(By.XPath(PREVIOUS_DAY));
            previousDay.Click();
            WaitForLoad();
        }
        public DateTime GetCalendarDay()
        {
            string dateText = WaitForElementIsVisible(By.XPath(CURRENT_CALENDAR_DAY)).Text;
            return DateTime.ParseExact(dateText.Substring(0, 10), "dd/MM/yyyy", CultureInfo.InvariantCulture);
        }

        public bool VerifierServiceExist(string serviceName)
        {
            var services = _webDriver.FindElements(By.XPath(SERVICES));
            foreach (var service in services) 
            { 
                if(service.Text.Equals(serviceName))
                { 
                    return true; 
                }
            }
            return false; 
        }

        public bool ErrorMessageExist(string errorMsg)
        {
            return WaitForElementIsVisible(By.XPath(ERROR_MSG)).Text.Equals(errorMsg);
        }
        public void MoveDateForward()
        {
            _webDriver.Navigate().Refresh();
            var moveForwardBtn = WaitForElementToBeClickable(By.XPath("//span[contains(@class,'dtRight')]"));
            moveForwardBtn.Click();           
            WaitForLoad();
        }

        public void MoveDateBackward()
        {
            _webDriver.Navigate().Refresh();
            IWebElement moveBackwardBtn;
            moveBackwardBtn = WaitForElementToBeClickable(By.XPath("//span[contains(@class,'dtLeft')]"));
            moveBackwardBtn.Click();
          
            WaitForLoad();
        }
        public void UpdateQunatitySetComment(string qty,string comment)
        {
            WaitPageLoading();
            WaitPageLoading();
            var inputQty = WaitForElementIsVisible(By.XPath(QTY_BY_DATE));
            inputQty.Clear();
            inputQty.SendKeys(qty);
            inputQty.Click();
            var commentInput = WaitForElementIsVisible(By.XPath(COMMENT_BTN));
            commentInput.Click();
            var commentText = WaitForElementIsVisible(By.Id(COMMENT_AREA));
            commentText.Clear();
            commentText.SendKeys(comment);
            var submit = WaitForElementIsVisible(By.XPath(SUBMIT_BTN));
            submit.Click();
        }
        public void Validate()
        {
            WaitForLoad();
            var validateBtn = WaitForElementIsVisible(By.Id(VALIDATE));
            validateBtn.Click();
        }
        public void GoToDate(DateTime date)
        {
            WaitPageLoading();
            while (GetCalendarDay().Date != date.Date)
            {
                if(GetCalendarDay().Date < date.Date)
                {
                    MoveDateForward();
                }
                if (GetCalendarDay().Date > date.Date)
                {
                    MoveDateBackward();
                }
            }
        }

        public string GetQuantityValue()
        {
            var qty = WaitForElementIsVisible(By.XPath(QTY));
            return qty.Text;
        }
        public string GetComment()
        {
            WaitForLoad();
            var comment = WaitForElementIsVisible(By.XPath(COMMENT_BTN));
            comment.Click();
            var value = WaitForElementIsVisible(By.XPath(COMMENT_VALUE));
            return value.Text;
        }

        public bool CheckIfModifiable()
        {

            if (isElementExists(By.XPath(QTY_INPUT)))
            {
                WaitForLoad();
                var inputQuantity = WaitForElementIsVisible(By.XPath(QTY_INPUT));
                var disabledAttribute = inputQuantity.GetAttribute("disabled");
                if (disabledAttribute == "true")
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }            
            else
            {
                return isElementExists(By.XPath(QTY_INPUT));
            }
        }

        public bool CheckIfModifiableByDay(string day, string service)
        {
            if (!isElementVisible(By.XPath(String.Format(QTY_INPUT_BY_DAY,service,day))))
            {
                return false; 
            }

            WaitForLoad();
            var inputQuantity = WaitForElementIsVisible(By.XPath(String.Format(QTY_INPUT_BY_DAY,service,day)));
            return inputQuantity.GetAttribute("disabled") != "true"; 
        }

        public List<string> CheckServiceNumber()
        {
            var lignes = _webDriver.FindElements(By.XPath(SERVICE_NUMBER));
            return lignes.Select(ligne => ligne.Text).ToList();
        }

    }
}
