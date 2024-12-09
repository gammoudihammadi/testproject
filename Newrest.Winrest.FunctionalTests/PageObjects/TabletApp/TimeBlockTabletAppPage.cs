using iText.Commons.Utils;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.TabletApp
{
    public class TimeBlockTabletAppPage : PageBase
    {

        //__________________________________Constantes_______________________________________

        private const string FLIGHT_TYPE_ICONE = "//star-image-svg";
        private const string FILTERS_BTN = "//button[@class=\"filter\"]";
        private const string SEARCH_INPUT = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[2]/input";
        private const string CUSTOMERS_INPUT = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[3]/ng-multiselect-dropdown/div/div[1]/span";
        private const string UNCHECK_ALL_CUSTOMERS = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[3]/ng-multiselect-dropdown/div/div[2]/ul[1]/li";
        private const string CUSTOMERS = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[3]/ng-multiselect-dropdown/div/div[2]/ul[2]/li[*]";
        private const string FLIGHT_TYPES_INPUT = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[9]/ng-multiselect-dropdown/div/div[1]/span";
        private const string UNCHECK_ALL_FLIGHT_TYPES = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[9]/ng-multiselect-dropdown/div/div[2]/ul[1]/li";
        private const string FLIGHT_TYPES = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[9]/ng-multiselect-dropdown/div/div[2]/ul[2]/li[*]";
        private const string CONFIRM_FILTER = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[10]/button";
        private const string ETD_FROM = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[4]/input";
        private const string ETD_TO = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[5]/input";
        private const string SHOW_MAJOR_FLIGHTS_ONLY_INPUT = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[6]/mat-slide-toggle/label/div/input";
        private const string SHOW_MAJOR_FLIGHTS_ONLY = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[6]/mat-slide-toggle/label/div";

        private const string HIDE_DONE_FLIGHTS_INPUT = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[7]/mat-slide-toggle/label/div/input";
        private const string HIDE_DONE_FLIGHTS = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[7]/mat-slide-toggle/label/div";

        private const string SHOW_LOAD_INPUT = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[8]/mat-slide-toggle/label/div/input";
        private const string SHOW_LOAD = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[8]/mat-slide-toggle/label/div";
        
        private const string DATE_FILTER = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[1]/div[2]";
        private const string YEAR = "/html/body/div[2]/div[2]/div/mat-datepicker-content/div[2]/mat-calendar/mat-calendar-header/div/div/button[1]/span[1]";
        private const string YEARS = "//mat-multi-year-view/table/tbody/tr[*]/td[*]/button/div[1]";
        private const string MONTHS = "//mat-year-view/table/tbody/tr[*]/td[*]/button/div[1]";
        private const string DAYS = "//mat-month-view/table/tbody/tr[*]/td[*]/button/div[1]";
        
        
        
        
        private const string FIRST_WORKSHOP_SATRT = "//virtual-scroller/div[2]/tbody/tr/td[10]/button[1]";
        private const string FIRST_WORKSHOP_END = "//virtual-scroller/div[2]/tbody/tr/td[10]/button[2]";

        private const string SECOND_WORKSHOP_START = "//virtual-scroller/div[2]/tbody/tr/td[11]/button[1]";
        private const string SECOND_WORKSHOP_END = "//virtual-scroller/div[2]/tbody/tr/td[11]/button[2]";

        private const string THIRD_WORKSHOP_START = "//virtual-scroller/div[2]/tbody/tr/td[12]/button[1]";
        private const string THIRD_WORKSHOP_END = "//virtual-scroller/div[2]/tbody/tr/td[12]/button[2]";


        private const string STATE = "//virtual-scroller/div[2]/tbody/tr/td[14]/cp-component-4button/div/div"; 
        private const string CUSTOMERS_NAMES = "//*/virtual-scroller/div[2]/tbody/tr[*]/td[2]/div/span[1]";
        private const string SHOW_HISTORY = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/table/virtual-scroller/div[2]/tbody/tr[1]/td[12]/div";
        private const string START_DATE_BTN = "/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/table/virtual-scroller/div[2]/tbody/tr/td[10]/button[1]";

        //__________________________________Variables_______________________________________


        public TimeBlockTabletAppPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        public enum FilterType
        {
            Search,
            Customers,
            EtdFrom,
            EtdTo,
            ShowMajorFlightsOnly,
            HideDoneFlights,
            ShowLoad,
            FlightType
        }

        public enum State
        {
            START,
            STARTED,
            DONE
        }
        public enum Color
        {
            Red,
            Green
        }
        public void Filter(FilterType filterType, object value)
        {
            Thread.Sleep(2000);
            var btn = WaitForElementExists(By.XPath(FILTERS_BTN));
            btn.Click();
         
            switch(filterType)
            {
                #region search
                case FilterType.Search:
                    var search = WaitForElementExists(By.XPath(SEARCH_INPUT));
                    search.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;
                #endregion
                #region customers
                case FilterType.Customers:
                    var customersInput = WaitForElementExists(By.XPath(CUSTOMERS_INPUT));
                    customersInput.Click();
                    var uncheckAllCustomers = WaitForElementExists(By.XPath(UNCHECK_ALL_CUSTOMERS));
                    uncheckAllCustomers.Click();
                    var customers = _webDriver.FindElements(By.XPath(CUSTOMERS));
                    foreach (var element in customers)
                    {
                        if (element.Text == value.ToString())
                            element.Click();
                    }
                    customersInput = WaitForElementExists(By.XPath(CUSTOMERS_INPUT));
                    customersInput.SendKeys(Keys.Enter);
                    break;
                #endregion
                #region etdfrom
                case FilterType.EtdFrom:
                    
                    var etdFrom = WaitForElementExists(By.XPath(ETD_FROM));
                    etdFrom.Click();

                        var hours = value.ToString().Substring(0, 2);
                        var minutes = value.ToString().Substring(2, 2);
                        var t = value.ToString().Substring(4,2); 
                    if(int.Parse(hours) < 0 || int.Parse(hours) >12)
                    {
                        throw new Exception("hours out of bounds");
                    }
                    etdFrom.SendKeys(hours);
                    if(int.Parse(minutes)<0 || int.Parse(minutes) > 60)
                    {
                        throw new Exception("minutes out of bounds");
                    }
                    etdFrom.SendKeys(minutes);
                    etdFrom.SendKeys(t);
                    break;
                #endregion
                #region etdto
                case FilterType.EtdTo:
                    var etdTo = WaitForElementExists(By.XPath(ETD_TO));
                    etdTo.Click();

                    hours = value.ToString().Substring(0, 2);
                    minutes = value.ToString().Substring(2, 2);
                    t = value.ToString().Substring(4, 2);
                    if (int.Parse(hours) < 0 || int.Parse(hours) > 12)
                    {
                        throw new Exception("hours out of bounds");
                    }
                    etdTo.SendKeys(hours);
                    if (int.Parse(minutes) < 0 || int.Parse(minutes) > 60)
                    {
                        throw new Exception("minutes out of bounds");
                    }
                    etdTo.SendKeys(minutes);
                    etdTo.SendKeys(t);
                    break;
                #endregion
                #region showmajorflightsonly
                case FilterType.ShowMajorFlightsOnly:
                    var showMajorCheckbox = WaitForElementExists(By.XPath(SHOW_MAJOR_FLIGHTS_ONLY_INPUT));
                    var checkboxToClick = WaitForElementExists(By.XPath(SHOW_MAJOR_FLIGHTS_ONLY));
                    if( (bool)value != showMajorCheckbox.Selected)
                    {
                        
                        checkboxToClick.Click();
                    }
                    break;
                #endregion
                #region hidedoneflights
                case FilterType.HideDoneFlights:
                    var hideDoneFlights = WaitForElementExists(By.XPath(HIDE_DONE_FLIGHTS_INPUT));
                    var hideDoneFlightsCheckBoxToClick = WaitForElementExists(By.XPath(HIDE_DONE_FLIGHTS));
                    if ((bool)value != hideDoneFlights.Selected)
                    {
                        hideDoneFlightsCheckBoxToClick.Click();
                    }
                    break;
                #endregion
                #region showload
                case FilterType.ShowLoad:
                    var showLoad = WaitForElementExists(By.XPath(SHOW_LOAD_INPUT));
                    var showLoadCheckBoxToClick = WaitForElementExists(By.XPath(SHOW_LOAD));
                    if ((bool)value != showLoad.Selected)
                    {
                        showLoadCheckBoxToClick.Click();
                    }
                    break;
                #endregion
                #region flighttype
                case FilterType.FlightType:
                    var flightTypesInput = WaitForElementExists(By.XPath(FLIGHT_TYPES_INPUT));
                    flightTypesInput.Click();
                    var uncheckAllFlightTypes = WaitForElementExists(By.XPath(UNCHECK_ALL_FLIGHT_TYPES));
                    uncheckAllFlightTypes.Click();
                    var flightTypes = _webDriver.FindElements(By.XPath(FLIGHT_TYPES));
                    foreach (var element in flightTypes)
                    {
                        if (element.Text == value.ToString())
                        {
                            element.Click();
                            break;
                        }
                    }
                    flightTypesInput = WaitForElementExists(By.XPath(FLIGHT_TYPES_INPUT));
                    flightTypesInput.Click();
                    flightTypesInput.Click();
                    break;
                #endregion
                #region default
                default: break;
                #endregion

            }
            var confirmBtn = WaitForElementExists(By.XPath(CONFIRM_FILTER));
            confirmBtn.Click();
            WaitPageLoading();
        }
        public bool VerifyFlightType(string color, string flightTypeName, bool filtrerFlightNo = false)
        {
            WaitForLoad();
            bool isFlightTypeFound = false;
            var stars = _webDriver.FindElements(By.XPath(FLIGHT_TYPE_ICONE));
            foreach (var elm in stars)
            {
                if (elm.GetAttribute("style").Contains(color) /*&& elm.GetAttribute("title").Contains(flightTypeName)*/)
                {
                    if (filtrerFlightNo)
                    {
                        // autre mise en page : flightNo en haut à aujourd'hui et les suivants J+1 J+2 J+3 sur ce même flightNo
                        return elm.GetAttribute("title").Contains(flightTypeName);
                    }
                    isFlightTypeFound = true;
                }
                else
                {
                    return false;
                }

            }

            //séparé de la première vérif et relance car erreur stale element reference: element is not attached to the page document
            stars = _webDriver.FindElements(By.XPath(FLIGHT_TYPE_ICONE));
            foreach (var elm in stars)
            {
                if (elm.GetAttribute("title").Contains(flightTypeName))
                {
                    isFlightTypeFound = true;
                }
                else
                {
                    return false;
                }
            }
            WaitForLoad();
            return isFlightTypeFound;

        }
        public void TimeBlockResetFilters()
        {
            var btnFilters = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[1]/div[2]/button"));
            btnFilters.Click();
            var searchInput = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[2]/input"));
            searchInput.ClearElement();
            var validateFilter = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[10]/button"));
            validateFilter.Click();
        }
        public void FilterByFlightType(string flightType)
        {
            Thread.Sleep(2500);
            var btnFilters = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[1]/div[2]/button"));
            btnFilters.Click();
            var input = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[9]/ng-multiselect-dropdown/div/div[1]/span"));
            input.Click();
            var uncheckall = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[9]/ng-multiselect-dropdown/div/div[2]/ul[1]/li/div"));
            uncheckall.Click();
            Thread.Sleep(2500);
            var flightTypes = _webDriver.FindElements(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[9]/ng-multiselect-dropdown/div/div[2]/ul[2]/li[*]"));
            foreach (var element in flightTypes)
            {
                if (element.Text == flightType)
                    element.Click();
            }
            var searchInput = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[2]/input"));
            searchInput.Click();
            var validateFilter = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[10]/button"));
            validateFilter.Click();
            Thread.Sleep(2500);
        }
        public void FilterByFlightNumber(string flightNumber)
        {
            Thread.Sleep(2500);
            var btnFilters = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[1]/div[2]/button"));
            btnFilters.Click();
            var searchInput = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[2]/input"));
            searchInput.SendKeys(flightNumber);
            var validateFilter = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[10]/button"));
            validateFilter.Click();
            WaitForLoad();
            WaitPageLoading();
        }
        public void FilterByETDFromTo(string from, string to)
        {
            Thread.Sleep(2500);
            var btnFilters = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[1]/div[2]/button"));
            btnFilters.Click();
            var fromInput = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[4]/input"));
            fromInput.SendKeys(from);
            var toInput = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[5]/input"));
            toInput.SendKeys(to);
            var validateFilter = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/div/div[10]/button"));
            validateFilter.Click();
        }
        public bool VerifyAll(string flightType, string color)
        {

            var stars = _webDriver.FindElements(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/table/virtual-scroller/div[2]/tbody/tr[*]/td[4]/div/star-image-svg"));
            foreach (var element in stars)
            {
                if (!element.GetAttribute("style").ToUpper().Contains(color.ToUpper()))
                {
                    return false;
                }
            }
            return true;
        }
        public bool VerifyEtd(string from, string to)
        {
            Thread.Sleep(1000);
            var elements = _webDriver.FindElements(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[2]/table/virtual-scroller/div[2]/tbody/tr[*]/td[9]"));

            foreach (var elm in elements)
            {
                if (Convert.ToInt32(elm.Text.Substring(0, 2)) < Convert.ToInt32(from) || Convert.ToInt32(elm.Text.Substring(0, 2)) > Convert.ToInt32(to))
                {
                    return false;
                }
            }
            return true;
        }

        public void SetDate(DateTime date)
        {
            CultureInfo ci = new CultureInfo("en-US");
            var jour = date.Day;
            var month = date.ToString("MMM", ci).ToUpper();
            var year = date.Year;
            var dateBtn = WaitForElementIsVisible(By.XPath(DATE_FILTER));
            dateBtn.Click();
            //set the year
            var yearBtn = WaitForElementIsVisible(By.XPath("/html/body/app-root/mat-sidenav-container/mat-sidenav-content/main/flight-workshop/div/div[1]/div[1]/mat-icon"));
            yearBtn.Click();
            Thread.Sleep(2000);
            WaitForLoad();
            SetYear(year);
            Thread.Sleep(2000);
            WaitForLoad();
            SetMonth(month);
            Thread.Sleep(2000);
            WaitForLoad();
            SetDay(jour);
            Thread.Sleep(2000);
            WaitHACCPHorizontalProgressBar();
        }

        public void WaitHACCPHorizontalProgressBar()
        {
            // attente de la progress bar
            int compteur = 1;
            bool vueSablier = false;
            while (compteur <= 1000)
            {
                try
                {
                    _webDriver.FindElement(By.ClassName("progress"));
                    vueSablier = true;
                    break;
                }
                catch
                {
                    compteur++;
                }
            }

            // attente de la fin de la progress bar
            compteur = 1;

            while (compteur <= 600 && vueSablier)
            {
                try
                {
                    var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(1));
                    wait.Until(ExpectedConditions.ElementIsVisible(By.ClassName("progress")));
                    compteur++;
                    // Ophélie : ajout d'un sleep pour augmenter le temps d'attente (équivalent à 1 minute max au total)
                    Thread.Sleep(100);
                }
                catch
                {
                    vueSablier = false;
                }
            }

            if (vueSablier)
            {
                throw new Exception("Délai d'attente dépassé pour le chargement de la page.");
            }
            WaitForLoad();
        }
        public string StartFirstWorkshop()
        {
            var firstWorkshopStart = WaitForElementIsVisible(By.XPath(FIRST_WORKSHOP_SATRT));
            firstWorkshopStart.Click();
            WaitPageLoading();
            firstWorkshopStart = WaitForElementIsVisible(By.XPath(FIRST_WORKSHOP_SATRT));
            if (firstWorkshopStart.GetAttribute("class").Contains("late"))
            {
                return Color.Red.ToString();
            }
            return Color.Green.ToString();
            
        }

        public string FirstWorkshopStarted()
        {
            var firstWorkshopStart = WaitForElementIsVisible(By.XPath(FIRST_WORKSHOP_SATRT));
            if (firstWorkshopStart.GetAttribute("class").Contains("late"))
            {
                return Color.Red.ToString();
            }
            return Color.Green.ToString();

        }


        public string EndFirstWorkshop()
        {
            var firstWorkshopEnd = WaitForElementExists(By.XPath(FIRST_WORKSHOP_END));
            firstWorkshopEnd.Click();
            WaitPageLoading();
            firstWorkshopEnd = WaitForElementExists(By.XPath(FIRST_WORKSHOP_END));
            if (firstWorkshopEnd.GetAttribute("class").Contains("late"))
            {
                return Color.Red.ToString();
            }
            return Color.Green.ToString();
        }

        public string FirstWorkshopEnded()
        {
            var firstWorkshopEnd = WaitForElementExists(By.XPath(FIRST_WORKSHOP_END));
            if (firstWorkshopEnd.GetAttribute("class").Contains("late"))
            {
                return Color.Red.ToString();
            }
            return Color.Green.ToString();
        }


        public string StartSecondWorkshop()
        {
            string color;
            var secondWorkshopStart = WaitForElementExists(By.XPath(SECOND_WORKSHOP_START));
            secondWorkshopStart.Click();
            WaitPageLoading();
            secondWorkshopStart = WaitForElementExists(By.XPath(SECOND_WORKSHOP_START));
            if (secondWorkshopStart.GetAttribute("class").Contains("late"))
            {
                color = Color.Red.ToString();
            }
            else
            {
                color = Color.Green.ToString();
            }
            WaitForLoad();
            return color;
        }
        public string SecondWorkshopStarted()
        {
            string color;
            var secondWorkshopStart = WaitForElementExists(By.XPath(SECOND_WORKSHOP_START));
            if (secondWorkshopStart.GetAttribute("class").Contains("late"))
            {
                color = Color.Red.ToString();
            }
            else
            {
                color = Color.Green.ToString();
            }
            WaitForLoad();
            return color;
        }

        public string EndSecondWorkshop()
        {
            string color;
            var secondWorkshopEnd = WaitForElementExists(By.XPath(SECOND_WORKSHOP_END));
            secondWorkshopEnd.Click();
            WaitPageLoading();
            secondWorkshopEnd = WaitForElementExists(By.XPath(SECOND_WORKSHOP_END));
            if (secondWorkshopEnd.GetAttribute("class").Contains("late"))
            {
                color = Color.Red.ToString();
            }
            else
            {
                color = Color.Green.ToString();
            }
            WaitForLoad();
            return color;
        }
        public string SecondWorkshopEnded()
        {
            string color;
            var secondWorkshopEnd = WaitForElementExists(By.XPath(SECOND_WORKSHOP_END));
            if (secondWorkshopEnd.GetAttribute("class").Contains("late"))
            {
                color = Color.Red.ToString();
            }
            else
            {
                color = Color.Green.ToString();
            }
            WaitForLoad();
            return color;
        }

        public string StartThirdWorkshop()
        {
            string color;
            var secondWorkshopStart = WaitForElementExists(By.XPath(THIRD_WORKSHOP_START));
            secondWorkshopStart.Click();
            WaitPageLoading();
            secondWorkshopStart = WaitForElementExists(By.XPath(THIRD_WORKSHOP_START));
            if (secondWorkshopStart.GetAttribute("class").Contains("late"))
            {
                color = Color.Red.ToString();
            }
            else
            {
                color = Color.Green.ToString();
            }
            WaitForLoad();
            return color;
        }
        public string EndThirdWorkshop()
        {
            string color;
            var secondWorkshopEnd = WaitForElementExists(By.XPath(THIRD_WORKSHOP_END));
            secondWorkshopEnd.Click();
            WaitPageLoading();
            secondWorkshopEnd = WaitForElementExists(By.XPath(THIRD_WORKSHOP_END));
            if (secondWorkshopEnd.GetAttribute("class").Contains("late"))
            {
                color = Color.Red.ToString();
            }
            else
            {
                color = Color.Green.ToString();
            }
            WaitForLoad();
            return color;
        }


        public bool VerifyState(State state)
        {
            _webDriver.Manage().Window.Maximize();
            var stateText = WaitForElementExists(By.XPath(STATE));
            if (state.ToString() == stateText.GetAttribute("innerText"))
            {
                return true;
            }
            return false;
        }

        public bool VerifyCustomers(string customer)
        {
            var customers = _webDriver.FindElements(By.XPath(CUSTOMERS_NAMES));
            List<string> customersList = new List<string>();
            foreach (var c in customers)
            {
                customersList.Add(c.Text);
            }
            for (int i = 0; i < customersList.Count; i++)
            {
                if (customersList[i] != customer)
                    return false;
            }
            return true;
        }
        // private methods

        private void SetYear(int year)
        {
            ReadOnlyCollection<IWebElement> years;
            years = _webDriver.FindElements(By.XPath(YEARS));
            foreach (var element in years)
            {
                if (int.Parse(element.Text) == year)
                {
                    element.Click();
                    break;
                }
            }
        }
        private void SetMonth(string month)
        {
            var currentMonth = WaitForElementIsVisible(By.XPath("//*[@id='mat-datepicker-0']/div/mat-month-view/table/tbody/tr[1]/td[1]"));
            var nbMonth = 0;
            while (nbMonth < 15 && currentMonth.Text.ToUpper() != month.ToUpper())
            {
                var nextMonthButtion = WaitForElementIsVisible(By.XPath("//*[@id='mat-datepicker-0']/mat-calendar-header/div/div/button[3]"));
                nextMonthButtion.Click();
                WaitForLoad();
                currentMonth = WaitForElementIsVisible(By.XPath("//*[@id='mat-datepicker-0']/div/mat-month-view/table/tbody/tr[1]/td[1]"));
                nbMonth++;
            }
        }
        private void SetDay(int day)
        {
            var element = _webDriver.FindElement(By.XPath("//*[@id=\"mat-datepicker-0\"]/div/mat-month-view/table/tbody/tr[*]/td[*]/button/span[contains(text(),' "+day+" ')]/parent::button"));
            element.Click();
        }

        protected new void WaitForLoad()
        {
            // pas de jQuery dans TabletApp
            var javaScriptExecutor = _webDriver as IJavaScriptExecutor;
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(61));

            Func<IWebDriver, bool> readyCondition = webDriver =>
                (bool)javaScriptExecutor.ExecuteScript("return (document.readyState == 'complete')");

            wait.Until(readyCondition);
        }
        public bool IsTimeBlockDisplayed()
        {
            WaitLoading();
            var timeBlockElement = _webDriver.FindElement(By.CssSelector("th.CustomerCodeCie"));
            return timeBlockElement != null && timeBlockElement.Displayed;
        }
        public bool IsPopUpshowHistoryVisible()
        {
            WaitLoading();
            IWebElement startDate = _webDriver.FindElement(By.XPath(START_DATE_BTN));
            startDate.Click();
            WaitLoading();
            IWebElement _showHistory = _webDriver.FindElement(By.XPath(SHOW_HISTORY));
            _showHistory.Click();

            IWebElement showHistoryPopUp = WaitForElementExists(By.Id("cdk-overlay-0"));
            if (showHistoryPopUp != null && showHistoryPopUp.Displayed)
            {
                return true;
            }
            else
                return showHistoryPopUp != null && showHistoryPopUp.Displayed;
        }
    }
}
