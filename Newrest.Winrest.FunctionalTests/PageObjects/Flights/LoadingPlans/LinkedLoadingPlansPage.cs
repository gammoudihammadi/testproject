using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newrest.Winrest.FunctionalTests.Utils;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans
{
    public class LinkedLoadingPlansPage : PageBase
    {

        public LinkedLoadingPlansPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes _________________________________________

        private const string GENERAL_INFORMATION = "hrefTabContentGeneral";
        private const string LOADING_PLAN_NUMBER = "//*[@id=\"LoadingPlanFormContainer\"]/div[1]/h1/span";

        // _________________________________________ Variables __________________________________________

        [FindsBy(How = How.Id, Using = GENERAL_INFORMATION)]
        private IWebElement _generalInformationPage;

        [FindsBy(How = How.XPath, Using = LOADING_PLAN_NUMBER)]
        private IWebElement _loadingPlanNumber;


        // _________________________________________ Méthodes ___________________________________________

        public int GetLinkedLoadingPlansNumber()
        {
            _loadingPlanNumber = WaitForElementIsVisible(By.XPath(LOADING_PLAN_NUMBER));
            return Convert.ToInt32(_loadingPlanNumber.Text);
        }

        public void ClickOnGeneralInformation()
        {
            _generalInformationPage = WaitForElementIsVisible(By.Id(GENERAL_INFORMATION));
            _generalInformationPage.Click();
            WaitForLoad();
        }
    }
}
