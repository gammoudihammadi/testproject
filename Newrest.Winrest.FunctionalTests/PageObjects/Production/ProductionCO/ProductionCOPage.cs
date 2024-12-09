using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.CustomerOrder;
using Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionManagement;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using System.Collections.ObjectModel;
using System.Globalization;
using System;
using System.Collections.Generic;
using System.Security.Policy;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.ProductionCO
{
    public class ProductionCOPage : PageBase
    {

        public ProductionCOPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________Constantes_____________________________________

        // Général
        private const string SEARCH_FILTER = "//*[@id=\"form-copy-from-tbSearchName\"]";
        private const string SITE = "//*[@id=\"drop-down-site\"]";
        private const string SEARCH = "//*[@id=\"searchServiceButton\"]";
        private const string Sites_FILTER = "SelectedSites_ms";
        private const string SEARCH_SITE = "/html/body/div[10]/div/div/label/input";
        private const string SITES_SELECT_ALL = "/html/body/div[10]/div/ul/li[1]/a/span[2]";
        private const string SITE_UNSELECT_ALL = "/html/body/div[10]/div/ul/li[2]/a/span[2]";

        // Tableau customer order
        private const string PLUS_BTN = "//*/button[text()='+']";
        private const string NEW_PRODUCTION_CO = "New Production CO";
        private const string LIST_ROWS = "//*[@id='tableListMenu']/tbody/tr[*]";
        private const string FIRST_ROW = "//*[@id=\"tabContentItemContainer\"]/table/tbody/tr[1]";
        

        //Filters
        private const string RESET_FILTER = "//*[@id=\"item-filter-form\"]/div[1]/a";
        private const string FILTER_DATE_FROM = "start-date-picker";
        private const string FILTER_DATE_TO = "end-date-picker";
        private const string FILTER_START_TIME_HOUR = "//*[@id=\"item-filter-form\"]/div[7]/div/input[2]";
        private const string FILTER_START_TIME_MIN = "//*[@id=\"item-filter-form\"]/div[7]/div/input[3]";
        private const string FILTER_END_TIME_HOUR = "//*[@id=\"item-filter-form\"]/div[8]/div/input[2]";
        private const string FILTER_END_TIME_MIN = "//*[@id=\"item-filter-form\"]/div[8]/div/input[3]";
        private const string SORTBY_FILTER = "cbSortBy";
        private const string SORTBY_PRODUCTIONDATE = "//*[@id=\"tabContentItemContainer\"]/table/tbody/tr[*]/td[6]";
        private const string SORTBY_NUMBER = "//*[@id=\"tabContentItemContainer\"]/table/tbody/tr[*]/td[2]";
        private const string FILTER_SEARCH = "SearchPattern";
        private const string ALL_SERVICES_NAME = "//*[@id=\"service-details-form\"]/table/tbody/tr[*]/td[1]";
        private const string ALLS_SERVICES_NAME = "//*[@id=\"tabContentItemContainer\"]/table/tbody/tr[3]/td[4]";
        private const string PRODUCTION_CO = "production-customerorderproduction-createbtn";
        private const string ALL_PRODUCTION_NUMBERS = "//*[@id=\"tabContentItemContainer\"]/table/tbody/tr[1]/td[2]";

        private const string PRINT_ICON = "//*[@id=\"production-customerorderproduction-showprint-1\"]";
        private const string LANGUAGE_LIST = "//*[@id=\"drop-down-language-from\"]";
        private const string PRINT_BTN = "//*[@id=\"submit-language\"]";
        private const string PRINTER_TOP_PAGE = "//*[@id=\"header-print-button\"]";
        private const string PDF_LIST = "/html/body/div[11]/div[2]/table/tbody/tr[*]";
        private const string STATUS = "/html/body/div[11]/div[2]/table/tbody/tr[2]/td[3]";
        private const string SYNC_BTN = "/html/body/div[11]/div[2]/div/a[3]";
        private const string PDF_ICON = "/html/body/div[11]/div[2]/table/tbody/tr[2]/td[4]/a";
        private const string DELETE_PRODUCTION_CO = "//*[@id=\"production-customerorderproduction-delete-1\"]";
        private const string CONFIRM_DELETE_PRODUCTION_CO = "//*[@id=\"dataConfirmOK\"]";
        private const string TELECHARGEMENT = "production-customerorderproduction-outputform-1";
        private const string START_TIME = "start-time-picker";
        private const string END_TIME = "end-time-picker";
        private const string DATE_ELEMENT = "//*[@id=\"tabContentItemContainer\"]/table/tbody/tr[1]/td[6]";




        //__________________________________Variables______________________________________

        // General



        // Tableau customer order
        [FindsBy(How = How.XPath, Using = PLUS_BTN)]
        private IWebElement _plusButton;

        [FindsBy(How = How.XPath, Using = NEW_PRODUCTION_CO)]
        private IWebElement _createNewProductionCO;

        [FindsBy(How = How.Id, Using = TELECHARGEMENT)]
        private IWebElement _telechargement;

        [FindsBy(How = How.Id, Using = SORTBY_FILTER)]
        private IWebElement _sortByBtn;

        [FindsBy(How = How.XPath, Using = SORTBY_PRODUCTIONDATE)]
        private IWebElement _sortByProductionDate;

        [FindsBy(How = How.ClassName, Using = "counter")]

        private IWebElement _totalNumber;

        [FindsBy(How = How.XPath, Using = SORTBY_NUMBER)]
        private IWebElement _sortByNumber;

        [FindsBy(How = How.XPath, Using = FIRST_ROW)]
        private IWebElement _firstRow;

        [FindsBy(How = How.XPath, Using = DELETE_PRODUCTION_CO)]
        private IWebElement _deleteButtonProductionCO;

        [FindsBy(How = How.XPath, Using = CONFIRM_DELETE_PRODUCTION_CO)]
        private IWebElement _confirmDeleteProductionCO;
        // ____________________________________________ Filtres ________________________________________________

        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = FILTER_DATE_FROM)]
        private IWebElement _dateFrom;

        [FindsBy(How = How.Id, Using = FILTER_DATE_TO)]
        private IWebElement _dateTo;

        [FindsBy(How = How.Id, Using = START_TIME)]
        public IWebElement _filterStartTimeHour;

        [FindsBy(How = How.Id, Using = END_TIME)]
        public IWebElement _filterEndTimeHour;

        [FindsBy(How = How.Id, Using = Sites_FILTER)]
        private IWebElement _SitesFilter;

        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSites;

        [FindsBy(How = How.XPath, Using = SITES_SELECT_ALL)]
        private IWebElement _siteSelectAll;

        [FindsBy(How = How.XPath, Using = SITE_UNSELECT_ALL)]
        private IWebElement _siteUnselectAll;

        [FindsBy(How = How.XPath, Using = SITE)]
        private IWebElement _site;
        [FindsBy(How = How.Id, Using = FILTER_SEARCH)]
        private IWebElement _searchFilter;

        [FindsBy(How = How.Id, Using = PRODUCTION_CO)]
        private IWebElement _productionCo;

        [FindsBy(How = How.XPath, Using = DATE_ELEMENT )]
        private IWebElement _dateElement;
        
        // ____________________________________________ Filtres ________________________________________________

        public enum FilterType
        {
            Search,
            DateFrom,
            DateTo,
            StartTime,
            EndTime,
            SortBy,
            Sites,
            Site
        }
        public void Filter(FilterType filterType, object value, string choice = null)
        {
            Actions action = new Actions(_webDriver);
            switch (filterType)
            {
                case FilterType.Search:
                    _searchFilter = WaitForElementIsVisible(By.Id(FILTER_SEARCH));
                    _searchFilter.SetValue(ControlType.TextBox, value);
                    _searchFilter.SendKeys(Keys.Tab);
                    break;
                case FilterType.DateFrom:
                    _dateFrom = WaitForElementIsVisible(By.Id(FILTER_DATE_FROM));
                    _dateFrom.SetValue(ControlType.DateTime,(DateTime)value);
                    _dateFrom.SendKeys(Keys.Tab);
                    break;

                case FilterType.DateTo:
                    _dateTo = WaitForElementIsVisible(By.Id(FILTER_DATE_TO));
                    _dateTo.SetValue(ControlType.DateTime, (DateTime)value);
                    _dateTo.SendKeys(Keys.Tab);
                    break;

                case FilterType.StartTime:
                    _filterStartTimeHour = WaitForElementIsVisible(By.Id(START_TIME));
                    _filterStartTimeHour.Clear();
                    _filterStartTimeHour.Click();
                    _filterStartTimeHour.SetValue(ControlType.TextBox, value.ToString().Substring(0, 2));
                    _filterStartTimeHour.SetValue(ControlType.TextBox, value.ToString().Substring(3, 2));
                    _filterStartTimeHour.SetValue(ControlType.TextBox, choice);
                    _filterStartTimeHour.SendKeys(Keys.Tab);
                    break;

                case FilterType.EndTime:
                    _filterEndTimeHour = WaitForElementIsVisible(By.Id(END_TIME));
                    _filterEndTimeHour.Clear();
                    _filterEndTimeHour.Click();
                    _filterEndTimeHour.SetValue(ControlType.TextBox, value.ToString().Substring(0, 2));
                    _filterEndTimeHour.SetValue(ControlType.TextBox, value.ToString().Substring(3, 2));
                    _filterEndTimeHour.SetValue(ControlType.TextBox, choice);
                    _filterEndTimeHour.SendKeys(Keys.Tab);

                    break;

                case FilterType.SortBy:
                    _sortByBtn = WaitForElementIsVisible(By.Id(SORTBY_FILTER));
                    _sortByBtn.Click();
                    var elementt = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortByBtn.SetValue(ControlType.DropDownList, elementt.Text);
                    _sortByBtn.Click();
                    break;

                case FilterType.Sites:
                    ScrollUntilElementIsInView(By.Id(Sites_FILTER));
                    if (isElementVisible(By.Id(Sites_FILTER)))
                    {
                        _SitesFilter = WaitForElementIsVisible(By.Id(Sites_FILTER));
                        _SitesFilter.Click();

                        if (value.Equals("ALL SITES"))
                        {
                            _searchSites = WaitForElementIsVisible(By.XPath(SEARCH_SITE));
                            _searchSites.SetValue(ControlType.TextBox, "");

                            _siteSelectAll = WaitForElementIsVisible(By.XPath(SITES_SELECT_ALL));
                            _siteSelectAll.Click();
                        }
                        else
                        {
                            _siteUnselectAll = WaitForElementIsVisible(By.XPath(SITE_UNSELECT_ALL));
                            _siteUnselectAll.Click();

                            _searchSites = WaitForElementIsVisible(By.XPath(SEARCH_SITE));
                            _searchSites.SetValue(ControlType.TextBox, value);
                            Thread.Sleep(1000);


                            if (isElementVisible(By.XPath("/html/body/div[10]/ul/li[*]/label/span[text()='" + value + " " + "-" + " " + value + "']")))
                            {
                                var valueToCheckCustomersType = WaitForElementIsVisible(By.XPath("//span[text()='" + value + " " + "-" + " " + value + "']")); 
                                    
                                valueToCheckCustomersType.SetValue(ControlType.CheckBox, true);

                            }
                        }

                        _SitesFilter.Click();

                    }
                    break;
                case FilterType.Site:
                    ComboBoxSelectById(new ComboBoxOptions(Sites_FILTER, (string)value));
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);
            }
            WaitForLoad();
            WaitPageLoading();
        }
        public bool IsCommentSubmittedSuccessfully()
        {
            var successMessageXPath = "//*[@id='tabContentItemContainer']/table/tbody/tr[1]/td[10]/span";
            var successMessageElement = _webDriver.FindElement(By.XPath(successMessageXPath));

            return successMessageElement.Displayed;
        }


        public DateTime GetDateFromXPath()
        {
            WaitForLoad();
            WaitPageLoading();
             _dateElement = WaitForElementIsVisible(By.XPath(DATE_ELEMENT));
            string dateText = _dateElement.Text;
            DateTime parsedDate;
           
            parsedDate = DateTime.ParseExact(dateText.Substring(0, 10), "dd/MM/yyyy",
                                              System.Globalization.CultureInfo.InvariantCulture);

            return parsedDate;
        }
        public ProductionCOCreateModalPage ProductionCOCreatePage()
        {
            ShowPlusMenuProductionCO();

            _createNewProductionCO = WaitForElementIsVisible(By.LinkText(NEW_PRODUCTION_CO));
            _createNewProductionCO.Click();
            WaitForLoad();

            return new ProductionCOCreateModalPage(_webDriver, _testContext);
        }

        public void ShowPlusMenuProductionCO()
        {
            _plusButton = WaitForElementIsVisible(By.XPath(PLUS_BTN));
            var actions = new Actions(_webDriver);
            actions.MoveToElement(_plusButton).Perform();
            WaitForLoad();
        }
      
        public bool IsSortedByProductionDate(string dateFormat)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            ReadOnlyCollection<IWebElement> dates;
            dates = _webDriver.FindElements(By.XPath(SORTBY_PRODUCTIONDATE));

            if (dates.Count == 0)
                return false;

            var ancientDate = DateTime.Parse(dates[0].Text, ci);

            foreach (var elm in dates)
            {
                try
                {
                    if (DateTime.Compare(ancientDate.Date, DateTime.Parse(elm.Text, ci)) < 0)
                        return false;

                    ancientDate = DateTime.Parse(elm.Text, ci);
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }
        public void ResetFilter()
        {
                _resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
                _resetFilter.Click();
                WaitForLoad();
            }
   
        public int GetCountResult()
        {
            WaitForLoad();
            _totalNumber = WaitForElementExists(By.ClassName("counter"));
            int nombre = Int32.Parse(_totalNumber.Text);
            return nombre;
        }

        public List<string> GetAllDates()
        {
            var rows = _webDriver.FindElements(By.XPath("//*[@id=\"tabContentItemContainer\"]/table/tbody/tr[*]/td[6]"));
            List<string> rowTexts = new List<string>();

            foreach (var row in rows)
            {
                string rowText = row.Text;

                rowTexts.Add(rowText);
            }
            return rowTexts;
        }


        public bool SortedByProductionDate (string dateFormat)
        {

            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");
            ReadOnlyCollection<IWebElement> dates;
            dates = _webDriver.FindElements(By.XPath(SORTBY_PRODUCTIONDATE));

            if (dates.Count == 0)
                return false;

            var ancientDate = DateTime.Parse(dates[0].Text, ci);

            foreach (var elm in dates)
            {
                try
                {
                    var x = DateTime.Parse(elm.Text, ci);
                    var y = ancientDate.Date; 
                    if (DateTime.Compare(DateTime.Parse(elm.Text, ci),ancientDate.Date)  < 0)
                        return false;

                    ancientDate = DateTime.Parse(elm.Text, ci);
                }
                catch
                {
                    return false;
                }
            }

            return true;
        }



        public bool IsSortedByNumber()
        {
            ReadOnlyCollection<IWebElement> elements = _webDriver.FindElements(By.XPath(SORTBY_NUMBER));

            List<int> intNumbers = new List<int>();

            int count = Math.Min(elements.Count, 3);

            for (int i = 0; i < count; i++)
            {
                string numberText = elements[i].Text.Trim();
                if (int.TryParse(numberText, out int result))
                {
                    intNumbers.Add(result);
                }
            }

            for (int i = 1; i < intNumbers.Count; i++)
            {
                if (intNumbers[i] < intNumbers[i - 1])
                {
                    return false;
                }
            }

            return true;
        }
        public ProductionCODetailsPage SelectFirstRow()
        {
            WaitForLoad();
            var firstRow = WaitForElementIsVisible(By.XPath(FIRST_ROW));
            firstRow.Click();
            WaitPageLoading();
            return new ProductionCODetailsPage(_webDriver, _testContext);
        }
       
        public bool CheckService(string targetService)
        {
            var allServices = _webDriver.FindElements(By.XPath(ALL_SERVICES_NAME));
            bool serviceFound = false;

            foreach (var service in allServices)
            {
                if (service.Text.Contains(targetService))
                {
                    serviceFound = true;
                    break;
                }
            }
            return serviceFound;
        }

        public bool CheckServices(string targetService)
        {
            var allServices = _webDriver.FindElements(By.XPath(ALLS_SERVICES_NAME));
            bool serviceFound = false;

            foreach (var service in allServices)
            {
                if (service.Text.Contains(targetService))
                {
                    serviceFound = true;
                    break;
                }
            }
            return serviceFound;
        }

        public bool AccessPage()
        {
            if (isElementExists(By.Id(PRODUCTION_CO)))
            {
                return true;
            }

            return false;
        }



        public void VerifyPrintIconExists()
        {
            WaitPageLoading();
            // Vérifie si l'icône d'impression est présente sur la ligne PCO
            var printIcon = _webDriver.FindElement(By.XPath(PRINT_ICON));
            if (printIcon == null)
            {
                throw new Exception("L'icône d'impression n'est pas disponible.");
            }
        }
        public void ClickPrintIcon()
        {
            WaitPageLoading();
            // Clique sur l'icône d'impression pour ouvrir la pop-up
            var printIcon = _webDriver.FindElement(By.XPath(PRINT_ICON));
            printIcon.Click();
        }
        public void SelectLanguageAndPrint(string language)
        {
            WaitPageLoading();
            // Sélectionne la langue dans la pop-up et clique sur "Print"
            var languageDropdown = new SelectElement(_webDriver.FindElement(By.XPath(LANGUAGE_LIST)));
            languageDropdown.SelectByText(language);

            var printButton = _webDriver.FindElement(By.XPath(PRINT_BTN));
            printButton.Click();
            WaitPageLoading();
        }
        public void ClickTopRightPrinterIcon()
        {
            WaitPageLoading();
            // Clique sur l'icône d'imprimante en haut à droite
            var topRightPrinterIcon = _webDriver.FindElement(By.XPath(PRINTER_TOP_PAGE));
            topRightPrinterIcon.Click();
            WaitPageLoading();
        }
        public string OpenLatestPDF()
        {
            // Localisez la liste des fichiers PDF
            var pdfList = _webDriver.FindElements(By.XPath(PDF_LIST));

            if (pdfList.Count <= 1)
            {
                throw new Exception("Aucun fichier PDF trouvé dans la liste.");
            }

            // Attendre que le statut du fichier soit 'finished'
            var lastPdfRow = pdfList[1];
            var statusElement = lastPdfRow.FindElement(By.XPath(STATUS));
            WaitPageLoading();
            // Cliquez sur l'icône de synchronisation jusqu'à ce que le statut soit 'finished'
            var syncButton = lastPdfRow.FindElement(By.XPath(SYNC_BTN));
            syncButton.Click();
            Thread.Sleep(3000);

            // Cliquez sur l'icône PDF pour ouvrir le fichier
            var pdfIcon = _webDriver.FindElement(By.XPath(PDF_ICON));
            pdfIcon.Click();

            // Attendez un moment pour que le PDF s'ouvre dans une nouvelle fenêtre
            WaitPageLoading();

            // Obtenez le titre de la nouvelle fenêtre
            var originalWindow = _webDriver.CurrentWindowHandle;
            var newWindow = _webDriver.WindowHandles.FirstOrDefault(handle => handle != originalWindow);

            if (newWindow == null)
            {
                throw new Exception("Aucune nouvelle fenêtre n'a été trouvée après l'ouverture du PDF.");
            }

            // Passez à la nouvelle fenêtre
            _webDriver.SwitchTo().Window(newWindow);
            WaitPageLoading();
            // Retourner l'URL du PDF pour l'utilisation dans l'assertion
            return _webDriver.Url;
        }
        public void deleteProductionCO()
        {
            _deleteButtonProductionCO = WaitForElementIsVisible(By.XPath(DELETE_PRODUCTION_CO));
            _deleteButtonProductionCO.Click();
            WaitPageLoading();

            _confirmDeleteProductionCO = WaitForElementIsVisible(By.XPath(CONFIRM_DELETE_PRODUCTION_CO));
            _confirmDeleteProductionCO.Click();
            WaitForLoad();
        }

        public bool IsCustomerOrderPresent(string name, string category, int quantity)
        {
            // Locate the rows in the search result table
            var rows = _webDriver.FindElements(By.XPath("//table[@class='table-class']//tr"));

            foreach (var row in rows)
            {
                var nameText = row.FindElement(By.XPath(".//td[1]")).Text;
                var categoryText = row.FindElement(By.XPath(".//td[2]")).Text;
                var quantityText = row.FindElement(By.XPath(".//td[3]")).Text;

                // Compare the details to ensure the order is not present
                if (nameText.Equals(name, StringComparison.OrdinalIgnoreCase) &&
                    categoryText.Equals(category, StringComparison.OrdinalIgnoreCase) &&
                    quantityText == quantity.ToString())
                {
                    return true;  // Order found, test will fail
                }
            }
            return false;  // Order not found, test will pass
        }

        public ProductionCOOutputFormPage GenerateNewOutputFrom()
        {
            _telechargement = _webDriver.FindElement(By.Id(TELECHARGEMENT));
            _telechargement.Click();
            WaitForLoad();
            return new ProductionCOOutputFormPage(_webDriver, _testContext);
        }
        public string GetFirstDateRow()
        {
            WaitForLoad();
            var firstDateRow = WaitForElementIsVisible(By.XPath("/html/body/div[2]/div/div[2]/div/table/tbody/tr[1]/td[6]"));
          
            return firstDateRow.Text;
        }
        public bool IsProdNumber(string prodNumber)
        {
            var allNumbers = _webDriver.FindElements(By.XPath(ALL_PRODUCTION_NUMBERS)); 
            bool prodNumberFound = false;

            foreach (var number in allNumbers)
            {
                if (number.Text.Contains(prodNumber))
                {
                    prodNumberFound = true;
                    break;
                }
            }
            return prodNumberFound;
        }
    }
}