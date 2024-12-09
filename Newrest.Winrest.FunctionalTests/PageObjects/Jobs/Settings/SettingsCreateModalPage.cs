using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium;
using Newrest.Winrest.FunctionalTests.Utils;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings
{
    public class SettingsCreateModalPage : PageBase
    {
        private const string NAME = "//*[@id=\"Name\"]";
        private const string FOLDER_NAME = "//*[@id=\"Folder\"]";
        private const string BNT_CREATE = "//*[@id=\"last\"]";
        private const string NAME_VALIDATION = "/html/body/div[3]/div/div/div/div/div/div/form/div[2]/div[1]/div/span";
        private const string FOLDER_VALIDATION = "/html/body/div[3]/div/div/div/div/div/div/form/div[2]/div[2]/div/span";

        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        [FindsBy(How = How.Id, Using = FOLDER_NAME)]
        private IWebElement _foldername;

        [FindsBy(How = How.Id, Using = BNT_CREATE)]
        private IWebElement _btn_create;


        [FindsBy(How = How.Id, Using = NAME_VALIDATION)]
        private IWebElement _nameValidation;


        [FindsBy(How = How.Id, Using = FOLDER_VALIDATION)]
        private IWebElement _folderValidation;


        public SettingsCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public SettingsPage FillField_CreateNewFileFlowProviders(string name, string foldername)
        {

            _name = WaitForElementIsVisible(By.XPath(NAME));
            _name.SetValue(ControlType.TextBox, name);

            _foldername = WaitForElementIsVisible(By.XPath(FOLDER_NAME));
            _foldername.SetValue(ControlType.TextBox, foldername);

            _btn_create = WaitForElementToBeClickable(By.XPath(BNT_CREATE));
            _btn_create.Click();
            WaitForLoad();
            return new SettingsPage(_webDriver, _testContext);
        }

        public bool FillField_CreateNewFileFlowProvidersWithoutData()
        {



            _btn_create = WaitForElementToBeClickable(By.XPath(BNT_CREATE));
            _btn_create.Click();
            WaitForLoad();

            var isNameVisible = isElementVisible(By.XPath(NAME_VALIDATION));

            var isFolderNameVisible = isElementVisible(By.XPath(FOLDER_VALIDATION));


            return isFolderNameVisible && isNameVisible;
        }
    }
}