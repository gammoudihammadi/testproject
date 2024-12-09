using DocumentFormat.OpenXml.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service
{
    public class ServiceLoadingPlanPage : PageBase
    {

        public ServiceLoadingPlanPage(IWebDriver _webDriver, TestContext _testContext) : base(_webDriver, _testContext)
        {
        }

        //________________________________________ Constantes _____________________________________________

        private const string LOADING_PLAN_NAME = "//*[@id=\"list-item-with-action\"]/div/div/div/div/table/tbody/tr/td[1]";
        private const string FIRST_ROW = "//*[@id=\"list-item-with-action\"]/div[1]/div/div/div/table/tbody/tr";


        //________________________________________ Variables ______________________________________________
        [FindsBy(How = How.XPath, Using = LOADING_PLAN_NAME)]
        private IWebElement _loadingPlanName;

        //________________________________________ Méthodes _______________________________________________

        public string GetLoadingPlanName()
        {
            WaitPageLoading();
            _loadingPlanName = WaitForElementIsVisible(By.XPath(LOADING_PLAN_NAME));
            return _loadingPlanName.Text;
        }
        public LoadingPlansGeneralInformationsPage SelectFirstLoadingPlan()
        {
            WaitForLoad();
            var firstRow = WaitForElementIsVisible(By.XPath(FIRST_ROW));
            firstRow.Click();
            return new LoadingPlansGeneralInformationsPage(_webDriver, _testContext);
        }
    }
}

