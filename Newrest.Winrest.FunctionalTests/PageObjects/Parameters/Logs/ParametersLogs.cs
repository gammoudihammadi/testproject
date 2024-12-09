using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Logs
{
    public class ParametersLogs : PageBase
    {
        


        public ParametersLogs(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        // ___________________________________ Constantes ______________________________________________
        private const string FILTER_DATE_FROM = "/html/body/div[2]/div/div[1]/div[1]/div/form/div/div[1]/div/div/input";
        private const string FILTER_DATE_TO = "/html/body/div[2]/div/div[1]/div[1]/div/form/div/div[2]/div/div/input";
        private const string AREA = "/html/body/div[2]/div/div[2]/div[2]/table/tbody/tr[*]/td[2]";
        private const string CONTROLLER = "/html/body/div[2]/div/div[2]/div[2]/table/tbody/tr[*]/td[3]";
        private const string INFO = "/html/body/div[2]/div/div[2]/div[2]/table/tbody/tr[*]/td[6]/span[2]";

        [FindsBy(How = How.XPath, Using = FILTER_DATE_FROM)]
        public IWebElement _dateFrom;

        [FindsBy(How = How.XPath, Using = FILTER_DATE_TO)]
        public IWebElement _dateTo;

        [FindsBy(How = How.XPath, Using = AREA)]
        public IWebElement _area;

        [FindsBy(How = How.XPath, Using = CONTROLLER)]
        public IWebElement _controller;

        [FindsBy(How = How.XPath, Using = INFO)]
        public IWebElement _info;

        // __________________________________ Méthodes __________________________________________________


        public enum FilterType
        {
            minDate, 
            maxDate
        }

        public void Filter(FilterType filterType, object value)
        {
            WaitForLoad();

            switch (filterType)
            {
                case FilterType.minDate:
                    _dateFrom = WaitForElementIsVisible(By.XPath(FILTER_DATE_FROM));
                    _dateFrom.SetValue(ControlType.DateTime, value);
                    break;

                case FilterType.maxDate:
                    _dateTo = WaitForElementIsVisible(By.XPath(FILTER_DATE_TO));
                    _dateTo.SetValue(ControlType.DateTime, value);
                    break;

                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }
            WaitPageLoading();
            WaitForLoad();
        }

        public bool VerifiedLogs(string Area, string Controller)
        {
   
            string error = "Error";
            for (int i = 0; i < 10 ; i++)
            {
                _area = WaitForElementIsVisible(By.XPath(AREA));
                WaitForLoad();
                _controller = WaitForElementIsVisible(By.XPath(CONTROLLER));
                WaitForLoad();
                _info = WaitForElementIsVisible(By.XPath(INFO));
                WaitForLoad();
                if ((_area.Text.ToLower() == Area.ToLower()) && (_controller.Text.ToLower() == Controller.ToLower()) && !(_info.Text.ToLower().Contains(error.ToLower())))
                {

                    return true;                 
                }

            }
            return false;
            

        }



    }
}

