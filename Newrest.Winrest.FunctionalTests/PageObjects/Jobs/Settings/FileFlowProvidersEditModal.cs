using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.Settings
{
    public  class FileFlowProvidersEditModal : PageBase
    {
        // ----------------------Constants------------------------------
        private const string CLOSE_X = "//*[@id=\"modal-1\"]/div/div/form/div[1]/button/span";
        private const string NAME = "Name";
        private const string FOLDER_NAME = "Folder";
        private const string SAVE = "last";
        private const string CLOSE = "//*[@id=\"modal-1\"]/div/div/form/div[3]/button[1]";
        private const string BNT_CREATE = "//*[@id=\"last\"]";
        private const string Plus_BTN_FILEFLOWPROVIDERS = "//*[@id=\"nav-tab\"]/li[8]/div/button";
        private const string Link_BTN_FILEFLOWProviders = "//*[@id=\"nav-tab\"]/li[8]/div/div/div/a";


        public FileFlowProvidersEditModal(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

            
        }
        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        [FindsBy(How = How.Id, Using = FOLDER_NAME)]
        private IWebElement _folderName;

        [FindsBy(How = How.Id, Using = SAVE)]
        private IWebElement _save;

        [FindsBy(How = How.XPath, Using = CLOSE)]
        private IWebElement _close;
        
        [FindsBy(How = How.XPath, Using = BNT_CREATE)]
        private IWebElement _btn_create;

        [FindsBy(How = How.XPath, Using = Plus_BTN_FILEFLOWPROVIDERS)]
        private IWebElement _plusbtnfileflowproviders;

        [FindsBy(How = How.XPath, Using = Link_BTN_FILEFLOWProviders)]
        private IWebElement _linkbtnfileflowproviders;

        public void SetName(string name)
        {
            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, name);
            WaitForLoad();
        }
        public void SetFolderName(string folderName)
        {
            _folderName = WaitForElementIsVisible(By.Id(FOLDER_NAME));
            _folderName.SetValue(ControlType.TextBox, folderName);
            WaitForLoad();
        }

        public void Save()
        {
            _save = WaitForElementIsVisible(By.Id(SAVE));
            _save.Click();
            WaitForLoad();
        }

        public void FillField_UpdateFileFlowProviders(string name, string updatefolderName)
        {
            var fileName = GetFileName();
            SetName(fileName + name);
            var foldername = GetFolderName();
            SetFolderName(foldername + updatefolderName);
            Save();
            WaitForLoad();

        }

        public FileFlowTypePage FillFields_UpdateFileFlow(string name, string updatefolderName)
        {
            SetName(name);
            SetFolderName(updatefolderName);
            Save();
            WaitForLoad();

            return new FileFlowTypePage(_webDriver, _testContext);

        }

        public string GetFileName()
        {
            var filename= WaitForElementIsVisible(By.Id(NAME));
            return filename.GetAttribute("value");
        }

        public string GetFolderName()
        {
            var foldername = WaitForElementIsVisible(By.Id(FOLDER_NAME));
            return foldername.GetAttribute("value");
        }

    }
}
