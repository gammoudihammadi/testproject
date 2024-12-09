using DocumentFormat.OpenXml.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.Flight;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionCO
{
    public class ProductionCODetailsPage : PageBase
    {
        public ProductionCODetailsPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        //__________________________________Constantes_______________________________________

        public const string BACK_TO_LIST = "/html/body/div[2]/a/span[1]";
        public const string FIRST_LINE_NUMBER = "/html/body/div[3]/div/div/div/div/form/div[1]/h4";
        public const string PRODUCTION_DATE = "prod-date-picker";
        public const string PRODUCTION_TIME = "ProductionTime";
        public const string SUBMIT_BTN = "/html/body/div[3]/div/div/div/div/form/div[3]/button[2]";
        public const string COMMENT = "Comment";
        public const string EXPEDITION_DATE = "//*[@id=\"edit-form\"]/div[2]/div[5]/div[2]/table/tbody/tr[1]/td[1]";
        public const string CLOSE = "//*[@id=\"edit-form\"]/div[1]/button/span";

        //__________________________________Variables_______________________________________

        [FindsBy(How = How.Id, Using = PRODUCTION_TIME)]
        private IWebElement _productiontime;
        [FindsBy(How = How.Id, Using = PRODUCTION_DATE)]
        private IWebElement _productiondate;
        [FindsBy(How = How.Id, Using = COMMENT)]
        private IWebElement _comment;
        [FindsBy(How = How.Id, Using = EXPEDITION_DATE)]
        private IWebElement _dateElement;
        [FindsBy(How = How.XPath, Using = CLOSE)]
        private IWebElement _close;


        public string GetProductionCONumber()
        {
            var firstLine = WaitForElementExists(By.XPath(FIRST_LINE_NUMBER));

            Match match = Regex.Match(firstLine.Text, @"\d+");

            if (match.Success)
            {
                string number = match.Value;
                return number;
            }
            else
            {
                return "0";
            }

        }
        public DateTime GetProductionDate()
        {
            var _productiondate = WaitForElementExists(By.Id(PRODUCTION_DATE));
            var dateText = _productiondate.GetAttribute("value");
            DateTime parsedDate;

            
            parsedDate = DateTime.ParseExact(dateText.Substring(0, 10), "dd/MM/yyyy",
                                              System.Globalization.CultureInfo.InvariantCulture);

            return parsedDate;
        }


        public void ApplyProductionTime(string productionTime,string choice)
        {
            _productiontime.SendKeys(productionTime);
            _productiontime.SetValue(ControlType.TextBox, choice);


            WaitPageLoading();
            var submit = WaitForElementIsVisible(By.XPath(SUBMIT_BTN));
            submit.Click();
            WaitPageLoading();
        }
        public DateTime getExpeditionDate()
        {
            IWebElement dateElement = _webDriver.FindElement(By.XPath(EXPEDITION_DATE));
            string dateText = dateElement.Text;
            DateTime parsedDate;
            parsedDate = DateTime.ParseExact(dateText.Substring(0, 10), "dd/MM/yyyy",
                                              System.Globalization.CultureInfo.InvariantCulture);

            return parsedDate;

        }
        public void AddComment(string comment)
        {
            var commentField = WaitForElementIsVisible(By.Id(COMMENT));
            commentField.SetValue(ControlType.TextBox, comment);
        }
        public List<DateTime> GetDatesFromXPathList()
        {
            List<DateTime> dateList = new List<DateTime>();
            IReadOnlyCollection<IWebElement> dateElements = _webDriver.FindElements(By.XPath("//*[@id='edit-form']/div[2]/div[5]/div[2]/table/tbody/tr[*]/td[1]"));

            foreach (IWebElement element in dateElements)
            {
                string dateText = element.Text;
                DateTime parsedDate;
                if (DateTime.TryParseExact(dateText, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                {
                    dateList.Add(parsedDate);
                }
                else
                {
                    throw new Exception($"Failed to parse the date from element with text: {dateText}");
                }
            }
            return dateList;
        }

        public void submit()
        {
            var submit = WaitForElementIsVisible(By.XPath(SUBMIT_BTN));
            submit.Click();
            WaitPageLoading();
        }

        public ProductionCOPage CloseViewDetails()
        {
            WaitPageLoading();
            _close = WaitForElementIsVisible(By.XPath(CLOSE));
            // légère animation (plus clair) lorsqu'on passe la sourie devant
            new Actions(_webDriver).MoveToElement(_close).Click().Perform();
            WaitForLoad();

            return new ProductionCOPage(_webDriver, _testContext);
        }

        public void ApplyProductionDate(DateTime productionDate)
        {
            _productiondate.SendKeys(productionDate.ToString());
            WaitPageLoading();
            var submit = WaitForElementIsVisible(By.XPath(SUBMIT_BTN));
            submit.Click();
            WaitPageLoading();
        }
    }
}
