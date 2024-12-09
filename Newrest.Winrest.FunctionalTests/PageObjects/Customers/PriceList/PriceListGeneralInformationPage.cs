using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Newrest.Winrest.FunctionalTests.Utils;
using System.Globalization;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Customers.PriceList
{
    public class PriceListGeneralInformationPage : PageBase
    {

        public PriceListGeneralInformationPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes _______________________________________

        // General
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";

        // Informations
        private const string NAME = "Name";
        private const string START_DATE = "StartDate";
        private const string END_DATE = "EndDate";

        //__________________________________ Variables ________________________________________

        // General
        [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;


        // Informations
        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        [FindsBy(How = How.Id, Using = START_DATE)]
        private IWebElement _startDate;

        [FindsBy(How = How.Id, Using = END_DATE)]
        private IWebElement _endDate;

        // _________________________________ Méthodes _________________________________________

        // General
        public PriceListPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new PriceListPage(_webDriver, _testContext);
        }

        // Informations
        public void ModifyGeneralInformations(string name, DateTime startDate, DateTime endDate)
        {
            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, name);

            _endDate = WaitForElementIsVisible(By.Id(END_DATE));
            _endDate.SetValue(ControlType.DateTime, endDate);

            _startDate = WaitForElementIsVisible(By.Id(START_DATE));
            _startDate.SetValue(ControlType.DateTime, startDate);

            // Temps d'enregistrement de la donnée
            WaitPageLoading();
            WaitForLoad();
        }

        public string GetPriceName()
        {
            _name = WaitForElementIsVisible(By.Id(NAME));
            return _name.GetAttribute("value");
        }

        public DateTime GetPriceStartDate(string dateFormat)
        {
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _startDate = WaitForElementIsVisible(By.Id(START_DATE));
            return DateTime.Parse(_startDate.GetAttribute("value"), ci).Date;
        }

        public DateTime GetPriceEndDate(string dateFormat)
        {
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _endDate = WaitForElementIsVisible(By.Id(END_DATE));
            return DateTime.Parse(_endDate.GetAttribute("value"), ci).Date;
        }

    }
}
