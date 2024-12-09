using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings
{
    public  class CreateSpecificJobNotificationModel : PageBase
    {
        private const string EMAIL = "Email";
        private const string CREATE_NEW= "last";
        private const string SPECIFIC = "SelectedSpecificScheduledJobIds_ms";
        private const string EMAIL_VALIDATORS = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[1]/div/span";


        [FindsBy(How = How.Id, Using = EMAIL)]
        private IWebElement _email;

        [FindsBy(How = How.Id, Using = SPECIFIC)]
        private IWebElement _specific;
        public CreateSpecificJobNotificationModel(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        public JobNotifications FillField_AddNewspecificjobnotification(string email, string Specific = null)
        {
            _email = WaitForElementIsVisible(By.Id(EMAIL));
            _email.SetValue(ControlType.TextBox, email);

            if (Specific != null)
            {
                ComboBoxSelectById(new ComboBoxOptions(SPECIFIC, (string)Specific, false));
            }

            var save = WaitForElementIsVisible(By.Id(CREATE_NEW));
            save.Click();
            WaitForLoad();
            return new JobNotifications(_webDriver, _testContext);
        }

        public bool VerifyEmailRequired()
        {

            return isElementVisible(By.XPath(EMAIL_VALIDATORS));

        }
    }
}
