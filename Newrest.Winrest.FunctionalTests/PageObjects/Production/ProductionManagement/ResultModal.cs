using DocumentFormat.OpenXml.ExtendedProperties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Customers.Service;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Datasheet;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionManagement
{
    public class ResultModal : PageBase
    {

        public ResultModal(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes _____________________________________

        private const string UNFOLD_ALL = "//a[contains(text(), 'Unfold All')]";
        private const string SERVICE_NAME = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td";
        private const string DATE_USE_CASE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div[1]/div/div/div[2]/table/tbody/tr/td[1]/span";
        private const string FLIGHT_USE_CASE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div[2]/div/table/tbody/tr[2]/td[2]";
        private const string DATE_DETAIL_USE_CASE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div[2]/div/table/tbody/tr[2]/td[3]";
        private const string CUSTOMER_USE_CASE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div[2]/div/table/tbody/tr[2]/td[4]";
        private const string SERVICE_USE_CASE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div[2]/div/table/tbody/tr[2]/td[5]";
        private const string SERVICE_QTE_USE_CASE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div[2]/div/table/tbody/tr[2]/td[6]";
        private const string RECIPE_USE_CASE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div[2]/div/table/tbody/tr[2]/td[7]";
        private const string NB_OCC_USE_CASE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div[2]/div/table/tbody/tr[2]/td[8]";
        private const string QTE_USE_CASE = "production-rawmaterialsitemusecase-itemquantity-1-1-1";
        private const string YIELD_USE_CASE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div[2]/div/table/tbody/tr[2]/td[10]";
        private const string METHODE_USE_CASE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div[2]/div/table/tbody/tr[2]/td[11]";
        private const string COEFFICIENT_USE_CASE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div[2]/div/table/tbody/tr[2]/td[12]";
        private const string QUANTITY_USE_CASE = "//*[@id=\"list-item-with-action\"]/div/div[2]/div/div[2]/div/table/tbody/tr[2]/td[13]";
        private const string EDIT_SERVICE = "//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[contains(text(),'{0}')]/a/span";
        private const string FLIGHT_NAME = "//*[starts-with(@id,\"content_usecase_\")]/div/table/tbody/tr[*]/td[2]";
        private const string CUSTOMER_NAME = "//*[starts-with(@id,\"content_usecase_\")]/div/table/tbody/tr[*]/td[4]";
        private const string EDIT_RECIPE = "//*[starts-with(@id,\"content_usecase_\")]/div/table/tbody/tr[{0}]/td[contains(text(),'{1}')]/../td[14]/a/span";
        private const string EDIT_DATASHEET = "//*[starts-with(@id,\"content_usecase_\")]/div/table/tbody/tr[{0}]/td[contains(text(),'{1}')]/../td[14]/a/span";
        private const string CLOSE = "//button[contains(text(), 'Close')]";
        private const string SITE_FILTRED = "SiteId";
        private const string DELIVERY_ROUND = "hrefTabContentItemContainer";
        private const string SELECTED_DELIVERIES = "SelectedDeliveriesIds_ms";
        private const string INPUT_SEARCH_DELIVERIE = "/html/body/div[12]/div/div/label/input";
        private const string INPUT_SEARCH_DELIVERIE_CHECKED = "//*[@id=\"ui-multiselect-2-SelectedDeliveriesIds-option-198\"]";
        private const string SELECTED_WORKSHOPS = "//*[@id=\"SelectedWorkshopsIds_ms\"]/span[2]";
        private const string INPUT_SEARCH_WORKSHOP = "/html/body/div[14]/div/div/label/input";
        private const string INPUT_SEARCH_WORKSHOP_CHECKED = "//*[@id=\"ui-multiselect-4-SelectedWorkshopsIds-option-4\"]";

        //__________________________________ Variables ______________________________________

        [FindsBy(How = How.XPath, Using = UNFOLD_ALL)]
        public IWebElement _unfoldAll;

        [FindsBy(How = How.XPath, Using = EDIT_SERVICE)]
        public IWebElement _editService;

        [FindsBy(How = How.XPath, Using = EDIT_RECIPE)]
        public IWebElement _editRecipe;

        [FindsBy(How = How.XPath, Using = EDIT_DATASHEET)]
        public IWebElement _editDatasheet;

        [FindsBy(How = How.XPath, Using = CLOSE)]
        public IWebElement _close;

        [FindsBy(How = How.XPath, Using = SERVICE_NAME)]
        public IWebElement _servicename;

        [FindsBy(How = How.XPath, Using = DATE_USE_CASE)]
        public IWebElement _dateusecase;

        [FindsBy(How = How.XPath, Using = FLIGHT_USE_CASE)]
        public IWebElement _flightusecase;

        [FindsBy(How = How.XPath, Using = DATE_DETAIL_USE_CASE)]
        public IWebElement _datedetailusecase;

        [FindsBy(How = How.XPath, Using = CUSTOMER_USE_CASE)]
        public IWebElement _customerusecase;

        [FindsBy(How = How.XPath, Using = SERVICE_USE_CASE)]
        public IWebElement _serviceusecase;

        [FindsBy(How = How.XPath, Using = SERVICE_QTE_USE_CASE)]
        public IWebElement _serviceqteusecase;

        [FindsBy(How = How.XPath, Using = RECIPE_USE_CASE)]
        public IWebElement _recipeusecase;

        [FindsBy(How = How.XPath, Using = NB_OCC_USE_CASE)]
        public IWebElement _nboocusecase;

        [FindsBy(How = How.XPath, Using = QTE_USE_CASE)]
        public IWebElement _qteusecase;

        [FindsBy(How = How.XPath, Using = YIELD_USE_CASE)]
        public IWebElement _yieldusecase;

        [FindsBy(How = How.XPath, Using = METHODE_USE_CASE)]
        public IWebElement _methodeusecase;

        [FindsBy(How = How.XPath, Using = COEFFICIENT_USE_CASE)]
        public IWebElement _cofficientusecase;

        [FindsBy(How = How.XPath, Using = QUANTITY_USE_CASE)]
        public IWebElement _quantityusecase;

        [FindsBy(How = How.Id, Using = SITE_FILTRED)]
        public IWebElement _siteFiltred;

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
                //servicePricePage.ClickOnGeneralInformationTab();


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
            if (isElementVisible(By.XPath(SERVICE_NAME)))
            {
                _webDriver.FindElement(By.XPath(SERVICE_NAME));
                return true;
            }
            else
            {
                return false;
            }
        }
        public string GetServiceName_UseCase()
        {
            _servicename = WaitForElementIsVisible(By.XPath(SERVICE_NAME));
            return _servicename.Text; ;
        }
        public string GetDate_UseCase()
        {
            _dateusecase = WaitForElementIsVisible(By.XPath(DATE_USE_CASE));
            string fullText = _dateusecase.Text;
            string date = fullText.Split(new[] { " - " }, StringSplitOptions.None)[0];
            return date;
        }


        public string GetFlight_UseCase()
        {
            _flightusecase = WaitForElementIsVisible(By.XPath(FLIGHT_USE_CASE));
            return _flightusecase.Text; ;
        }

        public string GetDate_Detail_UseCase()
        {
            _datedetailusecase = WaitForElementIsVisible(By.XPath(DATE_DETAIL_USE_CASE));
            return _datedetailusecase.Text; ;
        }
        public string GetCustomer_UseCase()
        {
            _customerusecase = WaitForElementIsVisible(By.XPath(CUSTOMER_USE_CASE));
            return _customerusecase.Text; ;
        }
        public string GetService_UseCase()
        {
            _serviceusecase = WaitForElementIsVisible(By.XPath(SERVICE_USE_CASE));
            return _serviceusecase.Text; ;
        }
        public string GetSERVICE_QTE_UseCase()
        {
            _serviceqteusecase = WaitForElementIsVisible(By.XPath(SERVICE_QTE_USE_CASE));
            return _serviceqteusecase.Text; ;
        }
        public string GetRecipe_UseCase()
        {
            _recipeusecase = WaitForElementIsVisible(By.XPath(RECIPE_USE_CASE));
            return _recipeusecase.Text; ;
        }
        public string GetNB_OCC_UseCase()
        {
            _nboocusecase = WaitForElementIsVisible(By.XPath(NB_OCC_USE_CASE));
            return _nboocusecase.Text; ;
        }
        public string GetQTEItem_UseCase()
        {
            if (IsDev())
            {
                _qteusecase = WaitForElementIsVisible(By.Id(QTE_USE_CASE));
            }
            else
            {
                _qteusecase = WaitForElementIsVisible(By.XPath("//*/div[@class='panel-body']/table/tbody/tr[2]/td[9]"));
            }
            return _qteusecase.Text; ;
        }
        public string GetYield_UseCase()
        {
            _yieldusecase = WaitForElementIsVisible(By.XPath(YIELD_USE_CASE));
            return _yieldusecase.Text; ;
        }
        public string GetMethode_UseCase()
        {
            _methodeusecase = WaitForElementIsVisible(By.XPath(METHODE_USE_CASE));
            return _methodeusecase.Text; ;
        }
        public string GetCoefficient_UseCase()
        {
            _cofficientusecase = WaitForElementIsVisible(By.XPath(COEFFICIENT_USE_CASE));
            return _cofficientusecase.Text; ;
        }
        public string GetQuantity_UseCase()
        {
            _quantityusecase = WaitForElementIsVisible(By.XPath(QUANTITY_USE_CASE));
            return _quantityusecase.Text; ;
        }

        public RecipeGeneralInformationPage EditRecipe(string recipeName, int ind)
        {
            WaitForLoad();
            _editRecipe = WaitForElementIsVisible(By.XPath(string.Format(EDIT_RECIPE, ind, recipeName)));

            _editRecipe.Click();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new RecipeGeneralInformationPage(_webDriver, _testContext);
        }

        public DatasheetDetailsPage EditDatasheet(string recipeName, int ind)
        {
            WaitLoading();

            //WaitForLoad();
            _editDatasheet = WaitForElementExists(By.XPath(string.Format(EDIT_DATASHEET, ind, recipeName)));
            Actions action = new Actions(_webDriver);
            action.MoveToElement(_editDatasheet).Perform();
            var element = WaitForElementToBeClickable(By.XPath(string.Format(EDIT_DATASHEET, ind, recipeName)));
            element.Click();
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

            _unfoldAll = WaitForElementIsVisible(By.XPath(UNFOLD_ALL));
            _unfoldAll.Click();
            WaitForElementIsVisible(By.XPath(FLIGHT_NAME));

            var flights = _webDriver.FindElements(By.XPath(FLIGHT_NAME));
            int indice = 2;
            RecipeGeneralInformationPage recipeGeneralInformationPage;

            for (var i = 0; i < flights.Count; i++)
            {
                action.MoveToElement(flights[i]).Perform();

                var datasheetPage = EditDatasheet(flights[i].Text.Trim(), indice+i);
                var recipePage = datasheetPage.EditRecipe();
                recipeGeneralInformationPage = recipePage.ClickOnGeneralInformationFromRecipeDatasheet();

                if (!recipeGeneralInformationPage.GetRecipeType().Equals(recipeType))
                {
                    return false;
                }
                //indice++;
                recipeGeneralInformationPage.Close();
            }

            return true;
        }

        public bool IsWorkshop(string workshop)
        {
            Actions action = new Actions(_webDriver);
            WaitLoading(); 
            _unfoldAll = WaitForElementIsVisible(By.XPath(UNFOLD_ALL));
            _unfoldAll.Click();
            WaitForElementIsVisible(By.XPath(FLIGHT_NAME));

            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> flights = _webDriver.FindElements(By.XPath(FLIGHT_NAME));
            RecipeGeneralInformationPage recipeGeneralInformationPage;

            for (int i = 0; i < flights.Count; i++)
            {
                action.MoveToElement(flights[i]).Perform();
                DatasheetDetailsPage datasheetPage = EditDatasheet(flights[i].Text.Trim(), 2+i);
                RecipesVariantPage recipePage = datasheetPage.EditRecipe();
                recipeGeneralInformationPage = recipePage.ClickOnGeneralInformationFromRecipeDatasheet();
                var workShop = recipeGeneralInformationPage.GetWorkshop();
                if (!recipeGeneralInformationPage.GetWorkshop().Equals(workshop))
                {
                    recipeGeneralInformationPage.Close();
                    return false;
                }
                recipeGeneralInformationPage.Close();
            }
            return true;
        }

        public bool IsCustomer(string customer)
        {
            _unfoldAll = WaitForElementIsVisible(By.XPath(UNFOLD_ALL));
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

            _unfoldAll = WaitForElementIsVisible(By.XPath(UNFOLD_ALL));
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
            WaitLoading();
            _unfoldAll = WaitForElementIsVisible(By.XPath(UNFOLD_ALL));
            _unfoldAll.Click();
            WaitForElementIsVisible(By.XPath(FLIGHT_NAME));
            var flights = _webDriver.FindElements(By.XPath(FLIGHT_NAME));
            RecipeGeneralInformationPage recipeGeneralInformationPage;
            for (var i = 0; i < flights.Count; i++)
            {
                action.MoveToElement(flights[i]).Perform();
                DatasheetDetailsPage datasheetPage = EditDatasheet(flights[i].Text.Trim(), 2);
                WaitForLoad();
                RecipesVariantPage recipePage = datasheetPage.EditFirstRecipe();//var recipePage = datasheetPage.EditRecipe();
                recipeGeneralInformationPage = recipePage.ClickOnGeneralInformationFromRecipeDatasheet();
                cookingModes.Add(recipeGeneralInformationPage.GetCookingMode());
                recipeGeneralInformationPage.Close();
            }
            return cookingModes;
        }

        public ResultPage CloseModal()
        {
            _close = WaitForElementIsVisible(By.XPath(CLOSE));
            _close.Click();
            WaitForLoad();

            return new ResultPage(_webDriver, _testContext);
        }

        public void UnfoldAll()
        {
            _unfoldAll = WaitForElementIsVisible(By.XPath(UNFOLD_ALL));
            _unfoldAll.Click();
        }

        public RecipeGeneralInformationPage FirstEditRecipe()
        {
            if (IsDev())
            {
                var datasheetPage = FirstEditDatasheet();
                var recipePage = datasheetPage.EditRecipe();
                var recipeGeneralInformationPage = recipePage.ClickOnGeneralInformationFromRecipeDatasheet();
            }
            else
            {
                Actions action = new Actions(_webDriver);
                WaitForLoad();
                var flightName = WaitForElementIsVisible(By.XPath(FLIGHT_NAME));
                flightName.Click();

                _editRecipe = WaitForElementIsVisible(By.XPath("//*/table/tbody/tr[2]/td[14]/a/span[contains(@class,'pencil')]"));
                action.MoveToElement(_editRecipe).Perform();

                _editRecipe.Click();
                WaitForLoad();

                //Results are opened in a new tab, switch the driver to the newly created one
                WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
                wait.Until((driver) => driver.WindowHandles.Count > 1);
                _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);
            }
            return new RecipeGeneralInformationPage(_webDriver, _testContext);
        }
        public DatasheetDetailsPage FirstEditDatasheet()
        {
            Actions action = new Actions(_webDriver);
            WaitForLoad();
            var flightName = WaitForElementIsVisible(By.XPath(FLIGHT_NAME));
            flightName.Click();

            var datasheetPage = EditDatasheet(flightName.Text.Trim(), 2);
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(30));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new DatasheetDetailsPage(_webDriver, _testContext);
        }
        public bool VerifyRecipePageOpened()
        {
            WaitForLoad();
            if (!isElementVisible(By.XPath("//*[@id=\"RecipeTabNav\"]/a")))
            {
                return false;
            }
            return true;
        }
        public bool IsWorkShop(string workshop)
        {
            Actions action = new Actions(_webDriver);
            WaitLoading();
            _unfoldAll = WaitForElementIsVisible(By.XPath(UNFOLD_ALL));
            _unfoldAll.Click();
            WaitForElementIsVisible(By.XPath(FLIGHT_NAME));

            System.Collections.ObjectModel.ReadOnlyCollection<IWebElement> flights = _webDriver.FindElements(By.XPath(FLIGHT_NAME));
            RecipeGeneralInformationPage recipeGeneralInformationPage;

            for (int i = 0; i < flights.Count; i++)
            {
                action.MoveToElement(flights[i]).Perform();
                DatasheetDetailsPage datasheetPage = EditDatasheet(flights[i].Text.Trim(), 2 + i);
                RecipesVariantPage recipePage = datasheetPage.EditRecipe();
                var workShops = recipePage.GetWorkShops();
                if (!workShops.Contains(workshop))
                {
                    recipePage.Close();
                    return false;
                }
                recipePage.Close();
            }
            return true;
        }
        public void ClickDeliveryRound()
        {
            var _deliveryRound = WaitForElementIsVisible(By.Id(DELIVERY_ROUND));
            _deliveryRound.Click();

        }
        public string GetSelectedSiteFilter()
        {
            WaitPageLoading();
            _siteFiltred = WaitForElementIsVisible(By.Id(SITE_FILTRED));
            var categorySelectElement = new SelectElement(_siteFiltred);

            return categorySelectElement.SelectedOption.Text;
        }
        public void ClickDeliveries()
        {
            var _deliverie = WaitForElementIsVisible(By.Id(SELECTED_DELIVERIES));
            _deliverie.Click();

        }
        public List<string> DeliveriesSelectedCombobox()
        {
            List<string> siblingNames = new List<string>();
            var deliveries = _webDriver.FindElements(By.XPath("//*[@name=\"multiselect_SelectedDeliveriesIds\"]"));

            foreach (var delivery in deliveries)
            {
                if (delivery.GetAttribute("checked") == "true")
                {
                    // Find the sibling (a <span> containing the name)
                    var sibling = delivery.FindElement(By.XPath("following-sibling::span"));
                    siblingNames.Add(sibling.Text.Trim());
                }

            }
            return siblingNames;
        }
        public bool VerifyDeliveries(string deliverie)
        {
            WaitPageLoading();
            ClickDeliveries();
            var _deliveriesFilterSearch = WaitForElementIsVisible(By.XPath(INPUT_SEARCH_DELIVERIE));
            _deliveriesFilterSearch.SetValue(ControlType.TextBox, deliverie);
            WaitPageLoading();
           return DeliveriesSelectedCombobox().Contains(deliverie);
        }
      
        public bool VerifyDeliveriesWrapper(string deliverie)
        {
            WaitPageLoading();
            return VerifyDeliveries(deliverie);
        }

        public void ClickWorkshops()
        {
            var _workshops = WaitForElementIsVisible(By.XPath(SELECTED_WORKSHOPS));
            _workshops.Click();

        }
        public bool VerifyWorkshops(string workshop1)
        {
            WaitPageLoading();
            ClickWorkshops();

            var _workshopsFilterSearch = WaitForElementIsVisible(By.XPath(INPUT_SEARCH_WORKSHOP));
            _workshopsFilterSearch.SetValue(ControlType.TextBox, workshop1);
            WaitPageLoading();
            var checkbox = WaitForElementExists(By.XPath(INPUT_SEARCH_WORKSHOP_CHECKED));
            bool isChecked = checkbox.Selected;

            return isChecked;
        }
        public bool VerifyWorkshopsWrapper(string workshop1)
        {
            WaitPageLoading();
            return VerifyWorkshops(workshop1);
        }

    }
}