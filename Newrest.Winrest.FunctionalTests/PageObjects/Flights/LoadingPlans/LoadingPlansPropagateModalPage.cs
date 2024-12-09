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
    public class LoadingPlansPropagateModalPage : PageBase
    {

        public LoadingPlansPropagateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver,
            testContext)
        {
        }


        // __________________________________________ Constantes ______________________________________________

        private const string SECURITY_CHECK = "loadingplan-propagate_securitychecks";
        private const string FIRST_SITE = "/html/body/div[4]/div/div/div[2]/div/form/table/tbody/tr[2]/td[1]/div";
        private const string PROPAGATE_BUTTON = "loadingplan-propagate";
        // __________________________________________ Variables _______________________________________________

        [FindsBy(How = How.XPath, Using = SECURITY_CHECK)]
        private IWebElement _securityCheck;

        [FindsBy(How = How.XPath, Using = FIRST_SITE)]
        private IWebElement _firstSite;

        [FindsBy(How = How.XPath, Using = PROPAGATE_BUTTON)]
        private IWebElement _propagateButton;

        //___________________________________________Pages___________________________________________________

        public string SecurityCheckLoadingPlan()
        {
            if (isElementVisible(By.XPath("/html/body/div[4]/div/div[@id=\"modal-1\" and @class=\"modal-content\"]")))
            {
             
                _firstSite = WaitForElementIsVisible(By.XPath(FIRST_SITE));
                _firstSite.Click();
                var site = WaitForElementIsVisible(By.XPath("//*[@id=\"form-SitesAvailable-CheckForSecurity\"]/table/tbody/tr[2]/td[2]/label"));
                var SiteSelected= site.Text;
                _securityCheck = WaitForElementIsVisible(By.Id(SECURITY_CHECK));
                _securityCheck.Click();
                WaitForLoad();
                WaitPageLoading();
                _propagateButton = WaitForElementIsVisible(By.Id(PROPAGATE_BUTTON));
                _propagateButton.Click();
                WaitForLoad();
                return SiteSelected;

            }else
            {
                return null;
            }
           
        }
    }
}
