using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans
{
    public class LoadingPlansSplitModalPage : PageBase
    {
        public LoadingPlansSplitModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // __________________________________________ Constantes ______________________________________________

        private const string NAME = "//*[@id=\"form-duplicate-loading-plan\"]/div/div[1]/div/input";
        private const string CONFIRM_SPLIT = "//*[@id=\"form-duplicate-loading-plan\"]/div/div[5]/button[2]";

        // __________________________________________ Variables _______________________________________________

        [FindsBy(How = How.XPath, Using = NAME)]
        private IWebElement _splitName;

        [FindsBy(How = How.XPath, Using = CONFIRM_SPLIT)]
        private IWebElement _splitBtn;

        //___________________________________________Pages___________________________________________________
        public LoadingPlansGeneralInformationsPage FillFields_SplitLoadingPlan(string loadingPlanNameSplit)
        {
            // Définition du nom du loading plan            
            _splitName = WaitForElementIsVisible(By.XPath(NAME));
            _splitName.SetValue(ControlType.TextBox, loadingPlanNameSplit);

            // Click sur le bouton "Split"
            _splitBtn = WaitForElementToBeClickable(By.XPath(CONFIRM_SPLIT));
            _splitBtn.Click();
            WaitForLoad();

            return new LoadingPlansGeneralInformationsPage(_webDriver, _testContext);
        }

    }
}