using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Production
{
    public class ParametersProductionModalPage : PageBase
    {

        //____________________________________________________Constantes_____________________________________________________________
        private const string VALUE_ALLOWED = "IsAllowed";
        private const string SAVE = "last";

        //____________________________________________________Variables_______________________________________________________________
        [FindsBy(How = How.Id, Using = VALUE_ALLOWED)]
        private IWebElement _valueAllowed;

        [FindsBy(How = How.Id, Using = SAVE)]
        private IWebElement _saveBtn;


        public ParametersProductionModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public bool SetAllowed()
        {
            _valueAllowed = WaitForElementExists(By.Id(VALUE_ALLOWED));

            var valueAllowed = _valueAllowed.GetAttribute("checked");

            if (valueAllowed == "true")
            {
                // Si la case était déjà cochée, on la laisse cochée pour les tests
                _saveBtn = WaitForElementIsVisible(By.Id(SAVE));
                _saveBtn.Click();
                WaitForLoad();
                return true;
            }
            else
            {
                // Si la case n'est pas cochée par défaut, on la coche pour les tests
                _valueAllowed.Click();
                _saveBtn = WaitForElementIsVisible(By.Id(SAVE));
                _saveBtn.Click();
                WaitForLoad();
                return false;
            }
        }

        public void RemoveAllowed()
        {
            _valueAllowed = WaitForElementExists(By.Id(VALUE_ALLOWED));

            var valueAllowed = _valueAllowed.GetAttribute("checked");

            if (valueAllowed == "true")
            {
                // Si la case était déjà cochée, on la décoche pour revenir à l'état initial
                _valueAllowed.Click();
                _saveBtn = WaitForElementIsVisible(By.Id(SAVE));
                _saveBtn.Click();
            }
            else
            {
                _saveBtn = WaitForElementIsVisible(By.Id(SAVE));
                _saveBtn.Click();
            }
        }

    }
}
