using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans
{
    public class LoadingPlansDuplicationModalPage : PageBase
    {
        public LoadingPlansDuplicationModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // _________________________________________ Constantes _____________________________________________

        private const string DUPLICATE_NAME = "//*[@id=\"form-duplicate-loading-plan\"]/div/div[1]/div/input";
        private const string DUPLICATE_START_DATE = "duplicate-start-date-picker";
        private const string DUPLICATE_END_DATE = "duplicate-end-date-picker";
        private const string DUPLICATE_BUTTON = "//*[@id=\"form-duplicate-loading-plan\"]/div/div[5]/button[2]";

        //___________________________________________ Variables _____________________________________________

        [FindsBy(How = How.XPath, Using = DUPLICATE_NAME)]
        private IWebElement _duplicateName;

        [FindsBy(How = How.Id, Using = DUPLICATE_START_DATE)]
        private IWebElement _duplicateStartDate;

        [FindsBy(How = How.Id, Using = DUPLICATE_END_DATE)]
        private IWebElement _duplicateEndDate;

        [FindsBy(How = How.XPath, Using = DUPLICATE_BUTTON)]
        private IWebElement _duplicateBtn;

        //___________________________________________ Pages _________________________________________________

        public LoadingPlansGeneralInformationsPage FillFields_DuplicateLoadingPlan(string loadingPlanNameBis, DateTime startDate, DateTime endDate)
        {
            // Définition du nom du loading plan            
            _duplicateName = WaitForElementIsVisible(By.XPath(DUPLICATE_NAME));
            _duplicateName.SetValue(ControlType.TextBox, loadingPlanNameBis);

            // Définition de l'attribut "StartDate"
            _duplicateStartDate = WaitForElementIsVisible(By.Id(DUPLICATE_START_DATE));
            _duplicateStartDate.SetValue(ControlType.DateTime, startDate);
            _duplicateStartDate.SendKeys(Keys.Enter);

            _duplicateEndDate = WaitForElementIsVisible(By.Id(DUPLICATE_END_DATE));
            _duplicateEndDate.SetValue(ControlType.DateTime, endDate);
            _duplicateEndDate.SendKeys(Keys.Enter);

            // Click sur le bouton "Duplicate"
            _duplicateBtn = WaitForElementToBeClickable(By.XPath(DUPLICATE_BUTTON));
            _duplicateBtn.Click();
            WaitForLoad();

            return new LoadingPlansGeneralInformationsPage(_webDriver, _testContext);

        }

    }
}