using DocumentFormat.OpenXml.Wordprocessing;
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

namespace Newrest.Winrest.FunctionalTests.PageObjects.Jobs.ScheduledJobs
{
    public class FileFlowsEditModal : PageBase
    {         // ----------------------Constants------------------------------
        private const string CLOSE_X = "//*[@id=\"modal-1\"]/div/div/form/div[1]/button/span";
        private const string NAME = "Name";
        private const string FOLDER_NAME = "Folder";
        private const string SAVE = "last";
        private const string COMPANY_CODE = "CompanyCode";
        private const string CONVERTER_TYPE = "ConverterType";
        public FileFlowsEditModal(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {

        }
    
        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;
        [FindsBy(How = How.Id, Using = FOLDER_NAME)]
        private IWebElement _folderName;
        [FindsBy(How = How.Id, Using = SAVE)]
        private IWebElement _save;
        [FindsBy(How = How.Id, Using = COMPANY_CODE)]
        private IWebElement _company_code;
        [FindsBy(How = How.Id, Using = CONVERTER_TYPE)]
        private IWebElement _converter_type;
        public void FillField_UpdateFileFlow(string name, string converterType, string companyCode)
        {
            SetName(name);

            _converter_type = WaitForElementIsVisible(By.Id(CONVERTER_TYPE));
            _converter_type.SetValue(PageBase.ControlType.DropDownList, converterType);
            WaitForLoad();

            _company_code = WaitForElementIsVisible(By.Id(COMPANY_CODE));
            _company_code.SetValue(PageBase.ControlType.TextBox, companyCode);
            WaitForLoad();

            Save();
            WaitForLoad();

        }
        public string GetFileName()
        {
            var filename = WaitForElementIsVisible(By.Id(NAME));
            return filename.GetAttribute("value");
        }
        public void SetName(string name)
        {
            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, name);
            WaitForLoad();
        }
         public void Save()
        {
            _save = WaitForElementIsVisible(By.Id(SAVE));
            _save.Click();
            WaitForLoad();
        }
    }
}
