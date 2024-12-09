using DocumentFormat.OpenXml.Bibliography;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using System;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Jobs
{
    public class CegidJobDetailPage : PageBase
    {
        public const string IMPORT_FILE_JOB_STATUS = "//*[@id=\"div-body\"]/div/div[1]/div[3]";

        public CegidJobDetailPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        public bool IsImportFileJobVisible()
        {
            return isElementVisible(By.XPath(IMPORT_FILE_JOB_STATUS));
        }

        public bool IsImportFileJobColored()
        {
            IWebElement element = _webDriver.FindElement(By.XPath(IMPORT_FILE_JOB_STATUS));
            if (element.GetAttribute("class").Contains("current")) return true;
            return false;
        }
    }
}
