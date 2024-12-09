using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Newrest.Winrest.FunctionalTests.PageObjects.Shared.PageBase;
using Newrest.Winrest.FunctionalTests.Utils;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Flights.LoadingPlans
{
    public class LoadingPlansHistoryModalPage : PageBase
    {
        public LoadingPlansHistoryModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
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

        public bool verifyHistoryModal()
        {
            WaitForLoad();
            if (isElementVisible(By.XPath("//div[@class=\"modal-content\"]/div[1]/h4[text()=\"History of loading plan modifications\"]")))
            {
                return true;

            }
            else
            {
                return false;
            }

        }

    }
}
