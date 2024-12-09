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

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings
{
    public class FileFlowProviderCreateModalPage : PageBase
    {
        private const string NAME = "Name";
        private const string FOLDER_NAME = "Folder";
        private const string BTN_CREATE = "last";
        private const string NAME_REQUIRED = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[1]/div/span";
        private const string FOLDERNAME_REQUIRED = "//*[@id=\"modal-1\"]/div/div/div/div/form/div[2]/div[2]/div/span";

        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        [FindsBy(How = How.Id, Using = FOLDER_NAME)]
        private IWebElement _foldername;

        [FindsBy(How = How.Id, Using = BTN_CREATE)]
        private IWebElement _btn_create;

        public FileFlowProviderCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public FileFlowProviderPage FillField_CreateNewFileFlowProvider(string name, string foldername)
        {
            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, name);

            _foldername = WaitForElementIsVisible(By.Id(FOLDER_NAME));
            _foldername.SetValue(ControlType.TextBox, foldername);

            _btn_create = WaitForElementToBeClickable(By.Id(BTN_CREATE));
            _btn_create.Click();
            WaitForLoad();
            return new FileFlowProviderPage(_webDriver, _testContext);
        }

        public bool IsNameRequiredMessageExist()
        {
            if (isElementVisible(By.XPath(NAME_REQUIRED))) return true;
            else return false;
        }
        public bool IsFolderNameRequiredMessageExist()
        {
            if (isElementVisible(By.XPath(FOLDERNAME_REQUIRED))) return true;
            else return false;

        }

    }

}
