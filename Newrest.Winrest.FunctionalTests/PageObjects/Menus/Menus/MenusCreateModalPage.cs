using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using Microsoft.VisualStudio.TestTools.UnitTesting;
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
using System.Windows.Interop;
using UglyToad.PdfPig.Outline;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Menus
{
    public class MenusCreateModalPage : PageBase
    {
        // ____________________________________________ Constantes ___________________________________________

        private const string NAME = "Name";
        private const string START_DATE = "StartDate";
        private const string END_DATE = "EndDate";
        private const string SITE = "dropdown-site";
        private const string CONCEPT = "dropdown-Concept";
        private const string ADD_VARIANT = "btn-add-variant";
        private const string SAVE_VARIANT = "last";
        private const string ADD_SERVICE = "btn-add-service";
        private const string SERVICE_FILTER = "SelectedServicesToMenu_ms";
        private const string UNCHECKALL_SERVICES = "/html/body/div[16]/div/ul/li[2]/a/span[2]";
        private const string SEARCH_SERVICE = "/html/body/div[16]/div/div/label/input";
        private const string SAVE_SERVICE = "last";
        private const string IS_ACTIVE = "check-box-isactive";
        private const string SAVE = "//*[@id=\"form-createdit-menu\"]/div[2]/div[3]/div/button[2]";
        private const string MSGS_VALIDATORS = "//*[@id=\"createFormMenu\"]/div[*]/div/span/span";
        private const string CALCULATION_METHOD = "dropdown-method-quantity";
        private const string VARIANTS_LIST = "//*[@id=\"menu-list-variant-popup\"]/div/div[1]/div/div[*]/label";

        // ____________________________________________ Variables ____________________________________________

        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        [FindsBy(How = How.Id, Using = START_DATE)]
        private IWebElement _startDate;

        [FindsBy(How = How.Id, Using = END_DATE)]
        private IWebElement _endDate;

        [FindsBy(How = How.Id, Using = SITE)]
        private IWebElement _site;
        
        [FindsBy(How = How.Id, Using = CONCEPT)]
        private IWebElement _concept;

        [FindsBy(How = How.Id, Using = ADD_VARIANT)]
        private IWebElement _addVariant;

        [FindsBy(How = How.Id, Using = SAVE_VARIANT)]
        private IWebElement _saveVariant;

        [FindsBy(How = How.Id, Using = ADD_SERVICE)]
        private IWebElement _addService;

        [FindsBy(How = How.Id, Using = SERVICE_FILTER)]
        private IWebElement _serviceFilter;

        [FindsBy(How = How.XPath, Using = UNCHECKALL_SERVICES)]
        private IWebElement _unCheckAllServices;

        [FindsBy(How = How.XPath, Using = SEARCH_SERVICE)]
        private IWebElement _searchService;

        [FindsBy(How = How.Id, Using = SAVE_SERVICE)]
        private IWebElement _saveService;

        [FindsBy(How = How.Id, Using = IS_ACTIVE)]
        private IWebElement _isActive;

        [FindsBy(How = How.XPath, Using = SAVE)]
        private IWebElement _saveBtn;

        [FindsBy(How = How.XPath, Using = MSGS_VALIDATORS)]
        private IWebElement _mesgs_validators;

        [FindsBy(How = How.Id, Using = CALCULATION_METHOD)]
        private IWebElement _calculMethod;
        // ____________________________________________ Méthodes _____________________________________________

        public MenusCreateModalPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        public MenusDayViewPage FillField_CreateNewMenu(string name, DateTime startDate, DateTime endDate, string site, string variant, string service = null, bool isActive = true, string methodCalculation="%")
        {
            WaitPageLoading();
            _name = WaitForElementIsVisibleNew(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, name);
            WaitForLoad();

            _startDate = WaitForElementIsVisibleNew(By.Id(START_DATE));
            _startDate.SetValue(ControlType.DateTime, startDate);
            _startDate.SendKeys(Keys.Tab);
            WaitForLoad();

            _endDate = WaitForElementIsVisibleNew(By.Id(END_DATE));
            _endDate.SetValue(ControlType.DateTime, endDate);
            _endDate.SendKeys(Keys.Tab);
            WaitForLoad();

            _site = WaitForElementIsVisibleNew(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site + " - " + site);
            WaitForLoad();


            // Ajout de la variante
                AddVariant(variant);
                WaitForLoad();

            // Définition du service
            if (service != null) {
                AddService(service);
                WaitForLoad();
            }

            _isActive = WaitForElementExists(By.Id(IS_ACTIVE));
            _isActive.SetValue(ControlType.CheckBox, isActive);
            WaitForLoad();

            _calculMethod = WaitForElementIsVisible(By.Id(CALCULATION_METHOD));
            _calculMethod.SetValue(ControlType.DropDownList, methodCalculation);
            WaitForLoad();

            // Click sur le bouton "Create"
            _saveBtn = WaitForElementToBeClickable(By.XPath(SAVE));
            _saveBtn.Click();
            WaitPageLoading();
            WaitForLoad();

            return new MenusDayViewPage(_webDriver, _testContext);
        }

        public void AddVariant(string variant)
        {
            Actions action = new Actions(_webDriver);

            // Définition de la variant
            _addVariant = WaitForElementIsVisibleNew(By.Id(ADD_VARIANT));
            _addVariant.Click();
            WaitForLoad();

            var selectedVariant = WaitForElementExists(By.XPath("//label[text()='" + variant + "']"));
            action.MoveToElement(selectedVariant).Perform();
            selectedVariant.SetValue(ControlType.CheckBox, true);
            WaitForLoad();

            _saveVariant = WaitForElementIsVisibleNew(By.Id(SAVE_VARIANT));
            _saveVariant.Click();
            WaitForLoad();

            // Temps de fermeture de la fenêtre
            WaitPageLoading();
            WaitForLoad();
        }

        public void AddService(string service)
        {
            Actions action = new Actions(_webDriver);

            _addService = WaitForElementIsVisibleNew(By.Id(ADD_SERVICE));
            _addService.Click();
            WaitForLoad();

            ComboBoxSelectById(new ComboBoxOptions(SERVICE_FILTER, service, false));

            _saveService = WaitForElementIsVisibleNew(By.Id(SAVE_SERVICE));
            _saveService.Click();

            // Temps de fermeture de la fenêtre
            WaitPageLoading();
        }

        public void AddSite(string site)
        {
            _site = WaitForElementIsVisible(By.Id(SITE));
            if (IsDev())
            {
                _site.SetValue(ControlType.DropDownList, site);
            }
            else
            {
                _site.SetValue(ControlType.DropDownList, site + " - " + site);
            }
            WaitForLoad();
        }

        public void ClickCreate()
        {
            // Click sur le bouton "Create"
            _saveBtn = WaitForElementToBeClickable(By.XPath(SAVE));
            _saveBtn.Click();
            WaitForLoad();
        }
        public bool IsMessagesValidatorsDisplayed()
        {
             var mesgs_validators = _webDriver.FindElements(By.XPath(MSGS_VALIDATORS));
            foreach (var validator in mesgs_validators)
            {
                if(!validator.Displayed)
                {
                    return false;
                }
            }
            return true;
        }

        public void FillField_CreateNewMenuNoVariant(string name, DateTime startDate, DateTime endDate, string site, string service = null, bool isActive = true)
        {
            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, name);
            WaitForLoad();

            _startDate = WaitForElementIsVisible(By.Id(START_DATE));
            _startDate.SetValue(ControlType.DateTime, startDate);
            _startDate.SendKeys(Keys.Tab);
            WaitForLoad();

            _endDate = WaitForElementIsVisible(By.Id(END_DATE));
            _endDate.SetValue(ControlType.DateTime, endDate);
            _endDate.SendKeys(Keys.Tab);
            WaitForLoad();

            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site + " - " + site);
            WaitForLoad();



            // Définition du service
            if (service != null)
            {
                AddService(service);
                WaitForLoad();
            }

            _isActive = WaitForElementExists(By.Id(IS_ACTIVE));
            _isActive.SetValue(ControlType.CheckBox, isActive);
            WaitForLoad();

            // Click sur le bouton "Create"
            _saveBtn = WaitForElementToBeClickable(By.XPath(SAVE));
            _saveBtn.Click();

            Thread.Sleep(2000);
        }

        public string GetAlertMessage()
        {
            var alert = _webDriver.SwitchTo().Alert();
            string messageError = alert.Text;
            _webDriver.SwitchTo().Alert().Accept();
            return messageError;
        }



        public List<string> GetVariantsOfSelectedSite(string site)
        {
            _site = WaitForElementIsVisible(By.Id(SITE));
            _site.SetValue(ControlType.DropDownList, site + " - " + site);
            WaitForLoad();

            Actions action = new Actions(_webDriver);

            // Définition de la variant
            _addVariant = WaitForElementIsVisible(By.Id(ADD_VARIANT));
            _addVariant.Click();
            WaitForLoad();

            var _listOfVariants = _webDriver.FindElements(By.XPath(VARIANTS_LIST));
            var variantsOfSelectedSite = _listOfVariants.Select(variant =>variant.Text).ToList();
            return variantsOfSelectedSite;

        }
        public int SelectDays()
        {
            int randomNumber = 0;
            Random random = new Random();
            randomNumber = random.Next(1, 8);
            Console.WriteLine(randomNumber);
            for (int i = 1; i <= 7; i++)
            {
                if(randomNumber != i)
                {
                    var _daySelected = WaitForElementExists(By.XPath("/html/body/div[3]/div/div/div/div/form/div[2]/div[1]/div/div[13]/div/ul/li[" + i.ToString() + "]/div/input"));
                    _daySelected.SetValue(ControlType.CheckBox, false);
                    WaitForLoad();
                }

            }
            return randomNumber;

        }
        public void SelectManyDays(List<string> listofdayweek)
        {           
           foreach (var day in listofdayweek)
           {
             var dayinweek = WaitForElementExists(By.XPath(String.Format("//*[@id=\"{0}\"]", day)));
             dayinweek.SetValue(ControlType.CheckBox, false);
           }
           WaitForLoad();     
        }
    }
}
