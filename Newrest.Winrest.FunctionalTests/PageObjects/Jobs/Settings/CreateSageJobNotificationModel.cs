using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings
{
    public class CreateSageJobNotificationModel : PageBase
    {
        private const string EMAIL_NEW_SAGE = "Email";
        private const string CREATE_NEW_SAGE = "last";
        private const string CEGID = "SelectedSageScheduledJobIds_ms";
        private const string EMAIL_VALIDATORS = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[1]/div/span";


        [FindsBy(How = How.Id, Using = EMAIL_NEW_SAGE)]
        private IWebElement _emailNewSage;

        [FindsBy(How = How.Id, Using = CEGID)]
        private IWebElement _cegid;
        public CreateSageJobNotificationModel(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }

        public JobNotifications FillField_AddNewsagejobnotification(string email, string Cegid=null)
        {
            _emailNewSage = WaitForElementIsVisible(By.Id(EMAIL_NEW_SAGE));
            _emailNewSage.SetValue(ControlType.TextBox, email);

            if (Cegid != null) 
            {
                ComboBoxSelectById(new ComboBoxOptions(CEGID, (string)Cegid, false));
            }

            var save = WaitForElementIsVisible(By.Id(CREATE_NEW_SAGE));
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
