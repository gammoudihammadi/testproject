using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.EarlyProduction
{
    public class EarlyProductionFavoriteModal : PageBase
    {

        public EarlyProductionFavoriteModal(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        private const string NAME = "Name";
        private const string SAVE = "last";
        

        public void Fill(string name)
        {
            var inputName = WaitForElementIsVisible(By.Id(NAME));
            inputName.SetValue(ControlType.TextBox, name);
            WaitForLoad();
        }

        public void Submit()
        {
            var save = WaitForElementIsVisible(By.Id(SAVE));
            save.Click();
            WaitForLoad();
        }
    }
}
