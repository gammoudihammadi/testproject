using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionManagement
{
    public class EditFavoriteModal : PageBase
    {
        public EditFavoriteModal(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes _____________________________________
        private const string INPUT_FAVORITE_NAME = "Name";
        private const string SUBMIT_FAVORITE_MODIF = "last";


        //__________________________________Variables______________________________________

        [FindsBy(How = How.Id, Using = INPUT_FAVORITE_NAME)]
        private IWebElement _nameFavorite;

        [FindsBy(How = How.Id, Using = SUBMIT_FAVORITE_MODIF)]
        private IWebElement _submitModif;

        //___________________________________ Méthodes _________________________________________

        public void RenameFavorite(string name)
        {
            Thread.Sleep(2500);
            WaitForLoad();
            _nameFavorite.SetValue(ControlType.TextBox, name);
        }

        public void SaveEdit()
        {
            _submitModif.Click();
            WaitForLoad();
        }
    }
}
