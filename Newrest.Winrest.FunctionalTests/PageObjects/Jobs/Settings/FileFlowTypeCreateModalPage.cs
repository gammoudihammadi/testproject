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
    public class FileFlowTypeCreateModalPage : PageBase
    {
        private const string NAME = "Name";
        private const string FOLDER_NAME = "Folder";
        private const string BNT_CREATE = "last";
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


        public FileFlowTypeCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public FileFlowTypePage FillField_CreateNewFileFlowType(string name, string foldername )
        {
          
            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, name);
            
            _foldername = WaitForElementIsVisible(By.Id(FOLDER_NAME));
            _foldername.SetValue(ControlType.TextBox, foldername);

            _btn_create = WaitForElementToBeClickable(By.Id(BNT_CREATE));
            _btn_create.Click();
            WaitForLoad();
            return new FileFlowTypePage(_webDriver, _testContext);
        }

        public void FillField_CreateNewFileFlowTypeWithoutData()
        {
            _btn_create = WaitForElementToBeClickable(By.Id(BNT_CREATE));
            _btn_create.Click();
            WaitForLoad();
        }

        public bool IsValidatorsExists()
        {
            var isNameVisible = isElementVisible(By.XPath(NAME_VALIDATION));
            var isFolderNameVisible = isElementVisible(By.XPath(FOLDER_VALIDATION));
            return isFolderNameVisible && isNameVisible;
        }

    }
}
