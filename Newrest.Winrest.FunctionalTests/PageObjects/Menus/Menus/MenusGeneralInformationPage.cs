using DocumentFormat.OpenXml.VariantTypes;
using FluentAssertions.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Menus
{
    public class MenusGeneralInformationPage : PageBase
    {
        // ___________________________________________ Constantes ___________________________________________

        // Général
        private const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        private const string DROPDOWN_BUTTON = "//div[@class='dropdown dropdown-print-button']";
        private const string EXPORT_TO_OTHER_MENU = "exportDataSheet";
        private const string DESTINATION_LINE = "tr-combo-line";
        private const string ADD_DESTINATION = "btn-create-new-row";
        private const string SITE_DESTINATION = "menudestination-site";
        private const string MENU_DESTINATION = "menudestination-menu";
        private const string CONFIRM_EXPORT = "last";
        private const string OK_BTN = "//*[@id=\"ExportSelectForm\"]/div[2]/button";

        private const string PRINT = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/a[2]";
        private const string CONFIRM_PRINT = "print";
        //private const string MENU_NAME = "Name";
        private const string MENU_NAME = "/html/body/div[3]/div/div[2]/div[3]/div/div/form/div/div[1]/div/div[1]/div/input";// "//*[@id=\"div-body\"]/div/h1";
        private const string CLEAR_SERVICES = "//*/a[text()='Clear Services']";
        private const string CLEAR_SERVICES_CONFIRM = "dataConfirmOK";
        private const string IS_ACTIVE = "check-box-isactive";
        private const string DELETE_ENTIRE_MENU = "//*[@id='div-delete-menu']/a";
        private const string DELETE_ENTIRE_MENU_CONFIRM = "dataConfirmOK";
        private const string ADD_SERVICES = "btn-add-service";

        // Informations
        private const string COMMERCIAL_NAME = "CommercialName1";
        private const string END_DATE = "EndDate";
        private const string BUDGET_MAX = "BudgetMax";
        private const string NUMBER_PAX = "TheoricalPaxNumber";
        private const string SALES_PRICE = "SalesPrice";
        private const string CLINIC = "IsClinic";
        private const string CALCULATION_METHOD = "dropdown-method-quantity";
        private const string CALCULATION_METHOD_SELECTED = "//*[@id=\"dropdown-method-quantity\"]/option[@selected = 'selected']";
        private const string CONCEPT_SELECTED = "//*[@id=\"dropdown-Concept\"]/option[2]";

        private const string PRINT_DEV = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div/div/a[3]";
        private const string PRINT_PATCH = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/a[2]";
        private const string LIST_SERVICES = "pListServiceOfMenu";
        private const string DEACTIVATE_INFO = "//*[@id=\"createFormMenu\"]/div[14]/div";
        //Modal Add Services
        private const string CATEGORY_SERVICES = "CategoryId";
        private const string SERVICES = "SelectedServicesToMenu_ms";
        private const string ADD_SERVICES_TOMENU = "last";
        private const string UNCHECK_ALL = "/html/body/div[12]/div/ul/li[2]/a";
        private const string CANCEL_SERVICES_MODAL = "/html/body/div[5]/div/div/form/div[2]/button[1]";
        private const string START_DATE_INPUT = "//*[@id=\"StartDate\"]";
        private const string END_DATE_INPUT = "//*[@id=\"EndDate\"]";
        private const string DATE_OF_WEEK = "//*[@id=\"createFormMenu\"]/div[13]/div/ul/li";
        private const string ISCLINICOPEN = "//*[@id=\"menusTab\"]/li[2]/a";
        private const string MENU_PLANNING_BY_DAY = "//*[starts-with(@id,\"MenuDay\")]/td/h4[contains(text(),'{0}')]";


        // ___________________________________________ Variables ____________________________________________

        // Général
        [FindsBy(How = How.XPath, Using = CONCEPT_SELECTED)]
        private IWebElement _conceptselected;

         [FindsBy(How = How.XPath, Using = BACK_TO_LIST)]
        private IWebElement _backToList;

        [FindsBy(How = How.Id, Using = DROPDOWN_BUTTON)]
        private IWebElement _dropdownButton;

        [FindsBy(How = How.Id, Using = EXPORT_TO_OTHER_MENU)]
        private IWebElement _exportToOtherMenu;

        [FindsBy(How = How.Id, Using = ADD_DESTINATION)]
        private IWebElement _addNewDestination;

        [FindsBy(How = How.Id, Using = SITE_DESTINATION)]
        private IWebElement _siteDestination;

        [FindsBy(How = How.Id, Using = MENU_DESTINATION)]
        private IWebElement _menuDestination;

        [FindsBy(How = How.Id, Using = CONFIRM_EXPORT)]
        private IWebElement _confirmExport;

        [FindsBy(How = How.XPath, Using = OK_BTN)]
        private IWebElement _okBtn;

        [FindsBy(How = How.XPath, Using = PRINT)]
        private IWebElement _print;

        [FindsBy(How = How.Id, Using = CONFIRM_PRINT)]
        private IWebElement _confirmPrint;

        [FindsBy(How = How.XPath, Using = CLEAR_SERVICES)]
        private IWebElement _clearServices;

        [FindsBy(How = How.Id, Using = CLEAR_SERVICES_CONFIRM)]
        private IWebElement _clearServicesConfirm;

        [FindsBy(How = How.Id, Using = IS_ACTIVE)]
        private IWebElement _isActive;

        [FindsBy(How = How.XPath, Using = DELETE_ENTIRE_MENU)]
        private IWebElement _deleteEntireMenu;

        [FindsBy(How = How.Id, Using = DELETE_ENTIRE_MENU_CONFIRM)]
        private IWebElement _deleteEntireMenuConfirm;

        // Informations
        [FindsBy(How = How.Id, Using = COMMERCIAL_NAME)]
        private IWebElement _commercialName;

        [FindsBy(How = How.Id, Using = END_DATE)]
        private IWebElement _endDate;

        [FindsBy(How = How.Id, Using = BUDGET_MAX)]
        private IWebElement _budgetMax;

        [FindsBy(How = How.Id, Using = NUMBER_PAX)]
        private IWebElement _numberPAX;

        [FindsBy(How = How.Id, Using = SALES_PRICE)]
        private IWebElement _salesPrice;

        [FindsBy(How = How.Id, Using = CLINIC)]
        private IWebElement _isClinic;

        [FindsBy(How = How.Id, Using = CALCULATION_METHOD)]
        private IWebElement _calculationMethod;

        [FindsBy(How = How.Id, Using = LIST_SERVICES)]
        private IWebElement _listServices;

        [FindsBy(How = How.Id, Using = ADD_SERVICES)]
        private IWebElement _addServices;

        [FindsBy(How = How.XPath, Using = UNCHECK_ALL)]
        private IWebElement _uncheckAll;


        [FindsBy(How = How.XPath, Using = MENU_NAME)]
        private IWebElement _menuname;

        [FindsBy(How = How.XPath, Using = ISCLINICOPEN)]
        private IWebElement _isClinicOpen;


        // ___________________________________________ Méthodes _____________________________________________

        public MenusGeneralInformationPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // Général
        public MenusPage BackToList()
        {
            _backToList = WaitForElementIsVisible(By.XPath(BACK_TO_LIST));
            _backToList.Click();
            WaitForLoad();

            return new MenusPage(_webDriver, _testContext);
        }

        public MenusDayViewPage ChooseMenuDay(string day)
        {
           var _day = WaitForElementIsVisible(By.XPath(String.Format(MENU_PLANNING_BY_DAY , day)));
            _day.Click();
            WaitForLoad();

            return new MenusDayViewPage(_webDriver, _testContext);
        }
        public void ExportToOtherMenu(string site, string menuName)
        {
            OpenDropdownButton();
            _exportToOtherMenu = WaitForElementIsVisible(By.Id(EXPORT_TO_OTHER_MENU));
            _exportToOtherMenu.Click();
            WaitForLoad();

            if (_webDriver.FindElement(By.Id(DESTINATION_LINE)).GetAttribute("class") != null
                && _webDriver.FindElement(By.Id(DESTINATION_LINE)).GetAttribute("class").Equals("hidden"))
            {
                _addNewDestination = WaitForElementIsVisible(By.Id(ADD_DESTINATION));
                _addNewDestination.Click();
                WaitForLoad();
            }

            _siteDestination = WaitForElementIsVisible(By.Id(SITE_DESTINATION));
            _siteDestination.SetValue(ControlType.DropDownList, site + " - " + site);

            _menuDestination = WaitForElementIsVisible(By.Id(MENU_DESTINATION));
            _menuDestination.SetValue(ControlType.DropDownList, menuName);

            _confirmExport = WaitForElementToBeClickable(By.Id(CONFIRM_EXPORT));
            _confirmExport.Click();
            WaitForLoad();
        }
        public void OpenDropdownButton()
        {
            _dropdownButton = WaitForElementIsVisible(By.XPath(DROPDOWN_BUTTON));
            _dropdownButton.Click();
            WaitForLoad();
        }

        public bool IsExportOK()
        {
            bool isOK = true;

            if (isElementVisible(By.XPath(OK_BTN)))
            {
                _okBtn = WaitForElementIsVisible(By.XPath(OK_BTN));
                _okBtn.Click();
                WaitForLoad();
                // Temps de fermeture de la fenêtre
                WaitPageLoading();
                WaitForLoad();
            }

            else
            {
                isOK = false;
            }

            return isOK;
        }

        public PrintReportPage Print(bool newVersionPrint)
        {
            OpenDropdownButton();
            var printD = WaitForElementIsVisible(By.XPath(PRINT_DEV));
            printD.Click();
            WaitForLoad();

            _confirmPrint = WaitForElementToBeClickable(By.Id(CONFIRM_PRINT));
            _confirmPrint.Click();
            WaitForLoad();

            if (newVersionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        // Informations

        public void SetCommercialName(string commercialName)
        {
            _commercialName = WaitForElementIsVisible(By.Id(COMMERCIAL_NAME));
            _commercialName.SetValue(ControlType.TextBox, commercialName);
        }

        public string GetCommercialName()
        {
            _commercialName = WaitForElementIsVisible(By.Id(COMMERCIAL_NAME));
            return _commercialName.GetAttribute("value");
        }

        public void SetEndDate(DateTime endDate)
        {
            _endDate = WaitForElementIsVisible(By.Id(END_DATE));
            _endDate.SetValue(ControlType.DateTime, endDate);
            WaitForLoad();

            //confirm action
            var confirmaction = WaitForElementIsVisible(By.Id("dataConfirmOK"));
            confirmaction.Click();
            WaitForLoad();
        }

        public string GetEndDate()
        {
            _endDate = WaitForElementIsVisible(By.Id(END_DATE));
            return _endDate.GetAttribute("value");
        }

        public void SetBudgetMax(string budget)
        {
            _budgetMax = WaitForElementIsVisible(By.Id(BUDGET_MAX));
            _budgetMax.SetValue(ControlType.TextBox, budget);

            WaitPageLoading();
        }

        public double GetBudgetMax(string decimalSeparator)
        {
            // Récupération du type de séparateur (, ou . selon les pays)
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _budgetMax = WaitForElementIsVisible(By.Id(BUDGET_MAX));
            return Double.Parse(_budgetMax.GetAttribute("value"), ci);
        }

        public void SetNumberPAX(string nbPAX)
        {
            Actions action = new Actions(_webDriver);

            _numberPAX = WaitForElementExists(By.Id(NUMBER_PAX));
            action.MoveToElement(_numberPAX).Perform();
            _numberPAX.ClearElement();
            _numberPAX.SendKeys(nbPAX);

            WaitPageLoading();
        }

        public double GetNumberPAX()
        {
            _numberPAX = WaitForElementExists(By.Id(NUMBER_PAX));
            return Double.Parse(_numberPAX.GetAttribute("value"));
        }

        public void SetSalesPrice(string salesPrice)
        {
            Actions action = new Actions(_webDriver);
            _salesPrice = WaitForElementExists(By.Id(SALES_PRICE));
            action.MoveToElement(_salesPrice).Perform();
            _salesPrice.SetValue(ControlType.TextBox, salesPrice);

            WaitPageLoading();
        }

        public double GetSalesPrice(string decimalSeparator)
        {
            // Récupération du type de séparateur (, ou . selon les pays)
            CultureInfo ci = decimalSeparator.Equals(",") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            _salesPrice = WaitForElementExists(By.Id(SALES_PRICE));
            return Double.Parse(_salesPrice.GetAttribute("value"), ci);
        }

        public void SetClinic(bool value)
        {
            Actions action = new Actions(_webDriver);
            _isClinic = WaitForElementExists(By.Id(CLINIC));
            action.MoveToElement(_isClinic).Perform();
            _isClinic.SetValue(ControlType.CheckBox, value);

            WaitPageLoading();
        }

        public bool IsClinic()
        {
            _isClinic = WaitForElementExists(By.Id(CLINIC));
            if (_isClinic.GetAttribute("checked") != null)
            {
                return true;
            }

            return false;
        }

        public void SetCalculationMethod(string method)
        {
            Actions action = new Actions(_webDriver);
            _calculationMethod = WaitForElementExists(By.Id(CALCULATION_METHOD));
            action.MoveToElement(_calculationMethod).Perform();
            _calculationMethod.SetValue(ControlType.DropDownList, method);

            WaitPageLoading();
        }

        public string GetCalculationMethod()
        {
            _calculationMethod = WaitForElementExists(By.XPath(CALCULATION_METHOD_SELECTED));
            return _calculationMethod.Text;
        }

        public MenusDayViewPage ClearService()
        {
            var html = _webDriver.FindElement(By.TagName("html"));
            html.SendKeys(Keys.PageDown);
            _clearServices = WaitForElementIsVisible(By.XPath(CLEAR_SERVICES));
            _clearServices.Click();
            if (isElementExists(By.Id(CLEAR_SERVICES_CONFIRM)))
            {
                _clearServicesConfirm = WaitForElementIsVisible(By.Id(CLEAR_SERVICES_CONFIRM));
                _clearServicesConfirm.Click();
            }

            WaitPageLoading();
            return new MenusDayViewPage(_webDriver, _testContext);
        }

        public void Activate()
        {
            var html = _webDriver.FindElement(By.TagName("html"));
            html.SendKeys(Keys.PageDown);
            _isActive = WaitForElementExists(By.Id(IS_ACTIVE));
            new Actions(_webDriver).MoveToElement(_isActive).Click().Perform();
            // animation : le bouton DeleteEntireMenu apparait
            WaitPageLoading();
            WaitForLoad();
        }

        public MenusPage DeleteEntireMenu()
        {
            _deleteEntireMenu = WaitForElementIsVisible(By.XPath(DELETE_ENTIRE_MENU));
            _deleteEntireMenu.Click();
            _deleteEntireMenuConfirm = WaitForElementIsVisible(By.Id(DELETE_ENTIRE_MENU_CONFIRM));
            _deleteEntireMenuConfirm.Click();
            WaitPageLoading();
            return new MenusPage(_webDriver, _testContext);
        }
        public List<string> GetServicesAddedToMenu()
        {
            //var services = new List<string>();
            _listServices = WaitForElementIsVisible(By.Id(LIST_SERVICES));
            if (_listServices.Text.Trim() == "No selected service")
            {
                return new List<string>();
            }
            var services = _listServices.Text.Trim().Split(',').ToList();
            return services.Select(s => s.Trim()).ToList();
        }
        public void AddServices(string category, List<string> services)
        {
            Actions action = new Actions(_webDriver);
            AddServicesToMenuModal();
            // u can uncomment this instruction to ensure the data (services Combo) is updated
            //DropdownListSelectById(CATEGORY_SERVICES, "CARNES");
            DropdownListSelectById(CATEGORY_SERVICES, category);

            if (isElementExists(By.Id(SERVICES)))
            {
                var servicesOptions = WaitForElementIsVisible(By.Id(SERVICES));
                servicesOptions.Click();
                WaitForLoad();
                _uncheckAll = WaitForElementIsVisible(By.XPath(UNCHECK_ALL));
                _uncheckAll.Click();
                foreach (var service in services)
                {
                    var selectedService = WaitForElementExists(By.XPath("/html/body/div[12]/ul/li[*]/label/span[text()='" + service + "']"));
                    action.MoveToElement(selectedService).Perform();
                    selectedService.SetValue(ControlType.CheckBox, true);
                }
                servicesOptions.Click();
                WaitForLoad();
            }
            var save = WaitForElementIsVisible(By.Id(ADD_SERVICES_TOMENU));
            save.Click();
            WaitPageLoading();
        }
        public void AddServicesToMenuModal()
        {
            _addServices = WaitForElementIsVisible(By.Id(ADD_SERVICES));
            _addServices.Click();
            WaitForLoad();
        }
        public List<string> GetAllServicesOptionsByCategory(string categoryService)
        {
            List<string> servicesOptions = new List<string>();
            AddServicesToMenuModal();
            DropdownListSelectById(CATEGORY_SERVICES, categoryService);
            WaitPageLoading();
            if (isElementExists(By.Id(SERVICES)))
            {
                var servicesCombobox = WaitForElementIsVisible(By.Id(SERVICES));
                servicesCombobox.Click();
                WaitForLoad();

                var services = _webDriver.FindElements(By.XPath("/html/body/div[12]/ul/li[*]/label/span"));
                foreach (var service in services)
                {
                    servicesOptions.Add(service.Text.Trim());
                }

                servicesCombobox.Click();
                WaitForLoad();
            }
            var close = WaitForElementIsVisible(By.XPath(CANCEL_SERVICES_MODAL));
            close.Click();
            WaitPageLoading();
            return servicesOptions;
        }

        public bool IsStartDateDisabled()
        {
            var _startDate = WaitForElementExists(By.XPath(START_DATE_INPUT));
            return _startDate.GetAttribute("disabled") == "true";
        }

        public bool IsEndDateDisabled()
        {
            var _startDate = WaitForElementExists(By.XPath(END_DATE_INPUT));
            return _startDate.GetAttribute("disabled") == "true";
        }

        public bool IsDateOfWeekDisabled()
        {
            var _datesOfWeek = _webDriver.FindElements(By.XPath(DATE_OF_WEEK));
            foreach (var dateOfWeek in _datesOfWeek)
            {
                if (dateOfWeek.GetAttribute("class") != "disabled") { return false; }
            }
            return true;
        }

        public void AddConcept()
        {
            _conceptselected = WaitForElementIsVisible(By.XPath(CONCEPT_SELECTED));
            _conceptselected.Click();
            Thread.Sleep(2000);
            WaitForLoad();
        }
        public string InfoDeactivate()
        {
            var info = WaitForElementIsVisible(By.XPath(DEACTIVATE_INFO));
            return info.Text;
        }
        public string GetMenuName()
        {
            if (isElementVisible(By.Id(MENU_NAME)))
            {
                _menuname = WaitForElementIsVisible(By.XPath(MENU_NAME));
                return _menuname.GetAttribute("value");
            }
            else
            {
                return "";
            }
        }
        
        public bool IsClinicOpen()
        {
            WaitPageLoading();
            return isElementExists(By.XPath(ISCLINICOPEN));
        }

    }
}
