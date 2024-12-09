using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing.Needs

{
    public class ResultModalNeeds : PageBase
    {

        public ResultModalNeeds(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes _____________________________________

        private const string UNFOLD_ALL = "//*[@id=\"modal-1\"]/div/div/div[2]/div/a";
        private const string SERVICE_NAME = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td";
        private const string EDIT_SERVICE = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[contains(text(),'{0}')]/a/span";
        private const string FLIGHT_NAME = "//*[starts-with(@id,\"content_usecase_\")]/div/table/tbody/tr[*]/td[2]";
        private const string CUSTOMER_NAME = "//*[starts-with(@id,\"content_usecase_\")]/div/table/tbody/tr[*]/td[4]";
        private const string EDIT_RECIPE = "//*[starts-with(@id,\"content_usecase_\")]/div/table/tbody/tr[*]/td[contains(text(),'{0}')]/../td[14]/a/span";
        private const string CLOSE = "//*[@id=\"modal-1\"]/div/div/div[4]/button";
        private const string GENERAL_INFORMATION = "//a[@id='hrefTabContentVariantDetailsInfos' and text()='General information']";



        //__________________________________ Variables ______________________________________

        [FindsBy(How = How.XPath, Using = UNFOLD_ALL)]
        public IWebElement _unfoldAll;

        [FindsBy(How = How.XPath, Using = EDIT_SERVICE)]
        public IWebElement _editService;

        [FindsBy(How = How.XPath, Using = EDIT_RECIPE)]
        public IWebElement _editRecipe;

        [FindsBy(How = How.XPath, Using = CLOSE)]
        public IWebElement _close;
        [FindsBy(How = How.XPath, Using = GENERAL_INFORMATION)]
        public IWebElement _general_Information;
        //___________________________________ Méthodes _________________________________________

        public ServicePricePage EditService(string serviceName)
        {
            _editService = WaitForElementIsVisible(By.XPath(string.Format(EDIT_SERVICE, serviceName)));
            _editService.Click();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new ServicePricePage(_webDriver, _testContext);
        }

        public bool IsServiceCategorie(string categorieService)
        {
            Actions action = new Actions(_webDriver);
            var services = _webDriver.FindElements(By.XPath(SERVICE_NAME));

            foreach (var elm in services)
            {
                action.MoveToElement(elm).Perform();
                var servicePricePage = EditService(elm.Text);

                WaitForLoad();
                var generalInformationPage = servicePricePage.ClickOnGeneralInformationTab();
                WaitForLoad();

                if (!generalInformationPage.GetCategory().Equals(categorieService))
                {
                    return false;
                }

                generalInformationPage.Close();
            }

            return true;
        }

        public bool IsService(string service)
        {
            var services = _webDriver.FindElements(By.XPath(SERVICE_NAME));

            foreach (var elm in services)
            {
                if (!elm.Text.Equals(service))
                {
                    return false;
                }
            }

            return true;
        }

        public bool IsServiceVisible()
        {
            try
            {
                _webDriver.FindElement(By.XPath(SERVICE_NAME));
                return true;
            }
            catch
            {
                return false;
            }
        }

        public DatasheetDetailsPage EditRecipe(string recipeName)
        {
            _editRecipe = WaitForElementExists(By.XPath(string.Format(EDIT_RECIPE, recipeName)));
            new Actions(_webDriver).MoveToElement(_editRecipe).Perform();
            _editRecipe.Click();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new DatasheetDetailsPage(_webDriver, _testContext);
        }

        public bool IsRecipeType(string recipeType)
        {
            Actions action = new Actions(_webDriver);
            _unfoldAll = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[2]/div/a"));
            _unfoldAll.Click();
            WaitForElementIsVisible(By.XPath(FLIGHT_NAME));

            var datasheetRecipes = _webDriver.FindElements(By.XPath(FLIGHT_NAME));

            foreach (var elm in datasheetRecipes)
            {
                action.MoveToElement(elm).Perform();
                var datasheetRecipePage = EditRecipe(elm.Text);
                datasheetRecipePage.EditRecipe();
                clickOnGeneraleInformation();
                if (!datasheetRecipePage.GetRecipeType().Equals(recipeType))
                {
                    return false;
                }

                datasheetRecipePage.Close();
            }

            return true;
        }

        public bool IsWorkshop(string workshop)
        {
            Actions action = new Actions(_webDriver);
                _unfoldAll = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[2]/div/a"));
            _unfoldAll.Click();
            WaitForElementIsVisible(By.XPath(FLIGHT_NAME));

            var datasheetRecipes = _webDriver.FindElements(By.XPath(FLIGHT_NAME));

            foreach (var elm in datasheetRecipes)
            {
                action.MoveToElement(elm).Perform();
                var datasheetRecipePage = EditRecipe(elm.Text);
                datasheetRecipePage.EditRecipe();
                clickOnGeneraleInformation();
                if (!datasheetRecipePage.GetWorkshop().Equals(workshop))
                {
                    return false;
                }

                datasheetRecipePage.Close();
            }

            return true;
        }

        public bool IsCustomer(string customer)
        {
                _unfoldAll = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[2]/div/a"));
            _unfoldAll.Click();
            WaitForElementIsVisible(By.XPath(CUSTOMER_NAME));

            var customers = _webDriver.FindElements(By.XPath(CUSTOMER_NAME));

            foreach (var elm in customers)
            {
                if (!elm.Text.Equals(customer))
                {
                    return false;
                }
            }

            return true;
        }

        public HashSet<string> GetFlights()
        {
            HashSet<string> flightList = new HashSet<string>();
                _unfoldAll = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[2]/div/a"));
            
            _unfoldAll.Click();
            WaitForElementIsVisible(By.XPath(FLIGHT_NAME));

            var flights = _webDriver.FindElements(By.XPath(FLIGHT_NAME));

            foreach (var elm in flights)
            {
                flightList.Add(elm.Text);
            }

            return flightList;
        }

        public HashSet<string> GetCookingModes()
        {
            Actions action = new Actions(_webDriver);
            HashSet<string> cookingModes = new HashSet<string>();
                _unfoldAll = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[2]/div/a"));
            _unfoldAll.Click();
            WaitForElementIsVisible(By.XPath(FLIGHT_NAME));

            var recipes = _webDriver.FindElements(By.XPath(FLIGHT_NAME));

            foreach (var elm in recipes)
            {
                action.MoveToElement(elm).Perform();

                var recipePage = EditRecipe(elm.Text);
                recipePage.EditRecipe();
                clickOnGeneraleInformation();
                cookingModes.Add(recipePage.GetCookingMode());
                recipePage.Close();
            }

            return cookingModes;
        }

        public ResultPageNeeds CloseModal()
        {
                _close = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[4]/button"));
            _close.Click();
            WaitForLoad();

            return new ResultPageNeeds(_webDriver, _testContext);
        }
        public void clickOnGeneraleInformation()
        {
            WaitForLoad();
            _general_Information = WaitForElementIsVisible(By.XPath(GENERAL_INFORMATION));
            _general_Information.Click();
        }

    }
}
