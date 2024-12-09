using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Text.RegularExpressions;
using System.Threading;
using UglyToad.PdfPig;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Purchasing
{
    public class SupplyOrderPage : PageBase
    {

        public SupplyOrderPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }
        // _______________________________________ Constantes __________________________________________________

        //Utilitaire 
        private const string EXTENDED_MENU = "//*[@id=\"div-body\"]/div/div[2]/div[1]/div/div[1]/button";
        private const string NEW_SUPPLY_ORDER = "New supply order";
        private const string FIRST_SO_NUMBER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr/td[2]";              
        private const string EXPORT = "btn-export-excel";
        private const string PRINT_RESULTS = "btn-print-supply-orders-report";
        private const string MODAL_PRINT_BUTTON = "printButton";
        private const string DELETE_FIRST_SO = "//*[@id=\"purchasing-supplyOrder-delete-1\"]";
        private const string CONFIRM_DELETE = "dataConfirmOK";
        private const string FIRST_SO_ITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]";
        private const string QTY_UNIT = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]/div[11]/span";
        private const string QTY_ITEM = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]/div[9]/span";

        

        // Filtres
        private const string RESET_FILTER_DEV = "ResetFilter";
        private const string SEARCH_FILTER = "SearchNumber";
        private const string SORTBY_FILTER = "cbSortBy";
        private const string SHOW_NOT_VALIDATED_FILTER = "ShowNotValidated";

        private const string SHOW_ALL_FILTER_DEV = "ShowAll";
        private const string SHOW_ACTIVE_FILTER_DEV = "ShowActiveOnly";
        private const string SHOW_NOT_ACTIVE_FILTER_DEV = "ShowInactiveOnly";

        private const string DATE_FROM_FILTER = "date-picker-start";
        private const string DATE_TO_FILTER = "date-picker-end";

        private const string VALIDATION_DATE_FILTER_DEV = "SearchByValidationDate";
        private const string VALIDATION_DATE_FILTER_PATCH = "//*[@value=\"ValidationDate\"]";

        private const string SITE_FILTER = "SelectedSites_ms";
        private const string UNSELECT_ALL = "/html/body/div[*]/div/ul/li[2]/a";
        private const string SEARCH_SITE = "/html/body/div[*]/div/div/label/input";

        private const string VERIFY_SITE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[3]";
        private const string VALIDATION = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/img[@alt='Valid']";
        private const string INACTIVE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[1]/img[@alt='Inactive']";
        private const string NUMBER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[2]";
        private const string NUMBER1 = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[2]";
        private const string STARTDATE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[5]";
        private const string START_DATE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[6]";
        private const string ENDDATE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[{0}]/td[6]";
        private const string END_DATE = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[6]";

        private const string FIRST_SUPPLY_ORDER = "//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[1]";
        private const string SEARCH_VALUE = "SearchNumber";
        private const string DOWNLOAD_ICON = "header-print-button";
        private const string DOWNLOAD_FILE = "/html/body/div[11]/div[2]/table/tbody/tr[2]/td[4]/a";
        private const string SUPPLY_ORDERS_NUMBERS = "/html/body/div[2]/div/div[2]/div[2]/table/tbody/tr[*]/td[2]";

        private const string SO_DELIVERY_LOCATION_SELECT = "SelectedSitePlaces_ms";
        private const string UNCHECK_ALL= "/html/body/div[11]/div/ul/li[2]/a";
        private const string DELIVERY_LOCATION_FILTER = "/html/body/div[11]/div/div/label/input";
        private const string SHOWITEMSNOTSUPPLIED = "ShowNotSupplied";
        private const string PRODQTY = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[*]/div/div/form/div[2]/div[9]/span";
        private const string SELECT_FIRST_SUPPLY_ORDER = "purchasing-supplyOrder-details-1";
        private const string FIRST_SUPPLY_ORDER_NOTSUPPLIED = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]";

        
        // ______________________________________ Variables ____________________________________________________

        [FindsBy(How = How.XPath, Using = EXTENDED_MENU)]
        private IWebElement _extendedMenu;
        [FindsBy(How = How.Id, Using = SELECT_FIRST_SUPPLY_ORDER)]
        private IWebElement _selectFirstSupplyOrder;

        [FindsBy(How = How.Id, Using = SHOWITEMSNOTSUPPLIED)]
        private IWebElement _showitemnotsupplied;

        [FindsBy(How = How.LinkText, Using = NEW_SUPPLY_ORDER)]
        private IWebElement _createNewSupplyOrder;

        [FindsBy(How = How.XPath, Using = FIRST_SO_NUMBER)]
        private IWebElement _firstSONumber;

        [FindsBy(How = How.XPath, Using = FIRST_SO_ITEM)]
        private IWebElement _firstSOItem;

        [FindsBy(How = How.Id, Using = EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.Id, Using = PRINT_RESULTS)]
        private IWebElement _printSupplyOrders;

        [FindsBy(How = How.Id, Using = MODAL_PRINT_BUTTON)]
        private IWebElement _modalPrintButton;

        [FindsBy(How = How.XPath, Using = DELETE_FIRST_SO)]
        private IWebElement _deleteFirstSO;

        [FindsBy(How = How.Id, Using = CONFIRM_DELETE)]
        private IWebElement _confirmDelete;

        // _________________________________________ Filtres ________________________________________________

        [FindsBy(How = How.Id, Using = RESET_FILTER_DEV)]
        private IWebElement _resetFilterDev;

        [FindsBy(How = How.Id, Using = SEARCH_FILTER)]
        private IWebElement _searchByNumber;

        [FindsBy(How = How.Id, Using = SORTBY_FILTER)]
        private IWebElement _sortBy;

        [FindsBy(How = How.Id, Using = SHOW_NOT_VALIDATED_FILTER)]
        private IWebElement _showNotValidated;

        [FindsBy(How = How.Id, Using = SHOW_ALL_FILTER_DEV)]
        private IWebElement _showAllDev;

        [FindsBy(How = How.Id, Using = SHOW_ACTIVE_FILTER_DEV)]
        private IWebElement _showActiveDev;

        [FindsBy(How = How.Id, Using = SHOW_NOT_ACTIVE_FILTER_DEV)]
        private IWebElement _showNotActiveDev;

        [FindsBy(How = How.Id, Using = DATE_FROM_FILTER)]
        private IWebElement _fromDate;

        [FindsBy(How = How.Id, Using = DATE_TO_FILTER)]
        private IWebElement _toDate;

        [FindsBy(How = How.Id, Using = VALIDATION_DATE_FILTER_DEV)]
        private IWebElement _byValidationDateDev;
        
        [FindsBy(How = How.XPath, Using = VALIDATION_DATE_FILTER_PATCH)]
        private IWebElement _byValidationDatePatch;

        [FindsBy(How = How.Id, Using = SITE_FILTER)]
        private IWebElement _siteFilter;

        [FindsBy(How = How.XPath, Using = UNSELECT_ALL)]
        private IWebElement _unselectAll;

        [FindsBy(How = How.XPath, Using = SEARCH_SITE)]
        private IWebElement _searchSite;


        public enum FilterType
        {
            ByNumber,
            SortBy,
            ShowNotValidated,
            ShowAll,
            ShowActive,
            ShowNotActive,
            DateFrom,
            DateTo,
            ByValidationDate,
            Site
        }

        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.ByNumber:
                    _searchByNumber = WaitForElementIsVisible(By.Id(SEARCH_FILTER));
                    _searchByNumber.SetValue(ControlType.TextBox, value);
                    break;
                case FilterType.SortBy:
                    _sortBy = WaitForElementIsVisible(By.Id(SORTBY_FILTER));
                    _sortBy.Click();
                    var element = WaitForElementIsVisible(By.XPath("//option[contains(@value,'" + value + "')]"));
                    _sortBy.SetValue(ControlType.DropDownList, element.Text);
                    _sortBy.Click();
                    break;
                case FilterType.ShowNotValidated:
                    _showNotValidated = WaitForElementExists(By.Id(SHOW_NOT_VALIDATED_FILTER));
                    action.MoveToElement(_showNotValidated).Perform();
                    _showNotValidated.SetValue(ControlType.CheckBox, value);
                    break;
                case FilterType.ShowAll:
                        _showAllDev = WaitForElementExists(By.Id(SHOW_ALL_FILTER_DEV));
                        action.MoveToElement(_showAllDev).Perform();
                        _showAllDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowActive:
                        _showActiveDev = WaitForElementExists(By.Id(SHOW_ACTIVE_FILTER_DEV));
                        action.MoveToElement(_showActiveDev).Perform();
                        _showActiveDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.ShowNotActive:
                        _showNotActiveDev = WaitForElementExists(By.Id(SHOW_NOT_ACTIVE_FILTER_DEV));
                        action.MoveToElement(_showNotActiveDev).Perform();
                        _showNotActiveDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.DateFrom:
                    _fromDate = WaitForElementIsVisible(By.Id(DATE_FROM_FILTER));
                    _fromDate.SetValue(ControlType.DateTime, value);
                    _fromDate.SendKeys(Keys.Tab);
                    break;
                case FilterType.DateTo:
                    _toDate = WaitForElementIsVisible(By.Id(DATE_TO_FILTER));
                    _toDate.SetValue(ControlType.DateTime, value);
                    _toDate.SendKeys(Keys.Tab);
                    break;
                case FilterType.ByValidationDate:
                        _byValidationDateDev = WaitForElementExists(By.Id(VALIDATION_DATE_FILTER_DEV));
                        action.MoveToElement(_byValidationDateDev).Perform();
                        _byValidationDateDev.SetValue(ControlType.RadioButton, value);
                    break;
                case FilterType.Site:
                    _siteFilter = WaitForElementIsVisible(By.Id(SITE_FILTER));
                    _siteFilter.Click();

                    _unselectAll = _webDriver.FindElement(By.XPath(UNSELECT_ALL));
                    _unselectAll.Click();
                    _searchSite = WaitForElementIsVisible(By.XPath(SEARCH_SITE));
                    _searchSite.SetValue(ControlType.TextBox, value + " - " + value);

                    var valueToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + " - " + value + "']"));
                    valueToCheck.SetValue(ControlType.CheckBox, true);
                    _siteFilter.Click();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }

            WaitPageLoading();
            WaitForLoad();
        }

        public void ResetFilter()
        {
            _resetFilterDev = WaitForElementIsVisible(By.Id(RESET_FILTER_DEV));
            _resetFilterDev.Click();
            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                Filter(FilterType.DateTo, DateUtils.Now);
            }
        }

        public Boolean VerifySite(string value)
        {
            bool valueBool = true;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

            var elements = _webDriver.FindElements(By.XPath(VERIFY_SITE));

            foreach (var elm in elements)
            {
                if (elm.Text != value)
                {
                    valueBool = false;
                }
            }

            return valueBool;
        }

        public bool CheckValidation(bool validated)
        {
            bool isValidated = true;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;


            for (int i = 0; i < tot; i++)
            {
                try
                {
                        _webDriver.FindElement(By.XPath(string.Format("//*[@id='list-item-with-action']/table/tbody/tr[{0}]/td[1]/div/i[contains(@class,'circle-check')]", i + 1)));
                    
                }
                catch
                {
                    isValidated = false;

                    if (validated)
                        return false;
                }
            }

            return isValidated;

        }

        public bool CheckStatus(bool active)
        {
            bool isActive = false;
            int tot = CheckTotalNumber() > 100 ? 100 : CheckTotalNumber();

            if (tot == 0)
                return false;

            for (int i = 0; i < tot; i++)
            {
                try
                {
                        _webDriver.FindElement(By.XPath(String.Format("//*[@id='list-item-with-action']/table/tbody/tr[1]/td[1]/div/i[contains(@class,'circle-xmark')]", i + 1)));
                   
                    if (active)
                        return false;
                }
                catch
                {
                    isActive = true;
                    if (!active)
                        return true;
                }
            }
            return isActive;
        }        

        public bool IsSortedByNumber()
        {
            bool valueBool = true;
            var ancientNumber = int.MaxValue;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }


            if (tot == 0)
                return false;

            var elements = _webDriver.FindElements(By.XPath(NUMBER1));

            foreach (var element in elements)
            {
                try
                {
                    int number = int.Parse(element.GetAttribute("innerText"));

                    if (number > ancientNumber)
                    { valueBool = false; }

                    ancientNumber = number;
                }
                catch 
                {
                    valueBool = false;
                }

            }
            return valueBool;
        }



        public bool IsSortedByStartDate()
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            var dateFormat = _fromDate.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            bool valueBool = true;
            DateTime ancientDate = DateTime.MinValue;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

            ReadOnlyCollection<IWebElement> elements;
            if (IsDev())
            {
                elements = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[6]"));
            }
            else {
                elements = _webDriver.FindElements(By.XPath(START_DATE));
            }

            foreach (var (element, index) in elements.Select((value, i) => (value, i)))
            {
                try
                {
                    string dateText = element.GetAttribute("innerText");
                    DateTime date = DateTime.Parse(dateText, ci);

                    if (index == 0)
                    {
                        ancientDate = date;
                    }

                    if (DateTime.Compare(ancientDate.Date, date) < 0)
                    {
                        valueBool = false;
                    }

                    ancientDate = date;
                }
                catch 
                {
                    valueBool = false;
                }
            }
            return valueBool;
        }

        public bool IsSortedByEndDate()
        {
            // Obtenir le format de date du champ de date
            var dateFormat = _toDate.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            // Flag pour vérifier si le tri est correct
            bool isSorted = true;
            DateTime previousDate = DateTime.MaxValue;  // On part d'une date maximale pour une comparaison décroissante

            // Récupérer tous les éléments de date dans la colonne correspondante
            ReadOnlyCollection<IWebElement> elements = _webDriver.FindElements(By.XPath("//*[@id='list-item-with-action']/table/tbody/tr/td[7]"));

            // Si aucun élément n'est trouvé, échouer immédiatement
            if (elements.Count == 0)
            {
                Console.WriteLine("Aucun élément de date trouvé dans la table.");
                return false;
            }

            foreach (var element in elements.Select((value, index) => (value, index)))
            {
                try
                {
                    // Extraire la date du texte
                    string dateText = element.value.GetAttribute("innerText").Trim();

                    // S'assurer que la chaîne de date n'est pas vide
                    if (string.IsNullOrEmpty(dateText))
                    {
                        Console.WriteLine($"La date à l'index {element.index} est vide.");
                        isSorted = false;
                        continue;
                    }

                    // Parse la date en utilisant le format spécifié
                    DateTime currentDate = DateTime.Parse(dateText, ci);

                    // Comparer les dates (tri décroissant)
                    if (element.index > 0 && currentDate > previousDate)
                    {
                        // Si l'ordre décroissant n'est pas respecté, on logge l'erreur
                        Console.WriteLine($"Erreur de tri à l'index {element.index} : Date précédente = {previousDate}, Date actuelle = {currentDate}");
                        isSorted = false;
                    }

                    previousDate = currentDate;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Erreur lors du parsing de la date à l'index {element.index}: {ex.Message}");
                    isSorted = false;
                }
            }

            return isSorted;
        }



        public Boolean IsFromDateRespected(DateTime FromDate)
        {
            var dateFormat = _fromDate.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            WaitForLoad();

            bool valueBool = true;
            int tot;
            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

            ReadOnlyCollection<IWebElement> elements;
            if (IsDev())
            {
                elements = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[6]"));
            }
            else
            {
                elements = _webDriver.FindElements(By.XPath(START_DATE));
            }
            foreach (var elem in elements)
            {
                try
                {
                    string dateText = elem.GetAttribute("innerText");
                    DateTime date = DateTime.Parse(dateText, ci);

                    if (FromDate.Date > date)
                    { valueBool = false; }
                }
                catch
                {
                    valueBool = false;
                }
            }
            return valueBool;
        }


        public Boolean IsToDateRespected(DateTime ToDate)
        {
            // Take the date format from the datepicker element and use it to format the date column to avoid date errors
            var dateFormat = _fromDate.GetAttribute("data-date-format");
            CultureInfo ci = dateFormat.Equals("dd/mm/yyyy") ? new CultureInfo("fr-FR") : new CultureInfo("en-US");

            bool valueBool = true;
            int tot;

            if (CheckTotalNumber() > 100)
            {
                tot = 100;
            }
            else
            {
                tot = CheckTotalNumber();
            }
            if (tot == 0)
                return false;

            // échéance
            ReadOnlyCollection<IWebElement> elements;
            if (IsDev())
            {
                elements = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[*]/td[6]"));
            }
            else
            {
                elements = _webDriver.FindElements(By.XPath(START_DATE));
            }
            foreach (var elem in elements)
            {
                try
                {
                    string dateText = elem.GetAttribute("innerText");
                    DateTime date = DateTime.Parse(dateText, ci);

                    if (ToDate.Date < date)
                    {
                        valueBool = false;
                    }
                }
                catch 
                {
                    valueBool = false;
                }

            }
            return valueBool;
        }

        // _________________________________________ Méthodes _______________________________________________



        public virtual void ShowExtendedMenu()
        {
            var actions = new Actions(_webDriver);
            _extendedMenu = WaitForElementIsVisible(By.XPath(EXTENDED_MENU));
            actions.MoveToElement(_extendedMenu).Perform();
        }

        public CreateSupplyOrderModalPage CreateNewSupplyOrder()
        {
            ShowPlusMenu();

            _createNewSupplyOrder = WaitForElementIsVisible(By.LinkText(NEW_SUPPLY_ORDER));
            _createNewSupplyOrder.Click();
            WaitForLoad();

            return new CreateSupplyOrderModalPage(_webDriver, _testContext);
        }

        public SupplyOrderPage SelectFirstSupplyOrder()
        {

            _selectFirstSupplyOrder = WaitForElementIsVisible(By.Id(SELECT_FIRST_SUPPLY_ORDER));
            _selectFirstSupplyOrder.Click();
            WaitForLoad();

            return new SupplyOrderPage(_webDriver, _testContext);
        }
        public string GetFirstSONumber()
        {
            _firstSONumber = WaitForElementIsVisible(By.XPath(FIRST_SO_NUMBER));
           return _firstSONumber.Text;
        }
        private const string XPathTemplate = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[{0}]/div/div/form/div[2]/div[9]/span";

         private const int IndexToSkip = 3;



        public bool GetQtyProd()
        {
            int index = 1; // Starting index for the rows in the table
            bool allQuantitiesValid = true;

            while (true) // Loop through each row until no more rows are found
            {
                try
                {
                    // XPath template pointing specifically to the "Prod. Qty" column
                    string xpath = $"/html/body/div[3]/div/div/div[2]/div[1]/div/form/div[2]/div/div/div[5]/div[2]/table/tbody/tr[{index}]/td[7]";
                    var element = _webDriver.FindElement(By.XPath(xpath));

                    if (element != null)
                    {
                        string text = element.Text.Trim();

                        // Extract numeric value from text (e.g., "10 KG" -> 10)
                        var match = System.Text.RegularExpressions.Regex.Match(text, @"\d+");
                        if (match.Success && int.TryParse(match.Value, out int quantity))
                        {
                            if (quantity != 0) // Vérifier si la quantité est différente de 0
                            {
                                allQuantitiesValid = false; // Si la quantité est différente de 0, ce n'est pas valide
                                break;
                            }
                        }
                        else
                        {
                            allQuantitiesValid = false; // En cas d'échec de l'extraction de la quantité
                            break;
                        }
                    }
                    else
                    {
                        break; // Arrêter la boucle si aucun élément trouvé pour cet index
                    }
                }
                catch (NoSuchElementException)
                {
                    break; // Arrêter la boucle s'il n'y a plus de lignes dans la table
                }

                index++;
            }

            return allQuantitiesValid;
        }

        public bool GetQtyProdAfterShowItem()
        {
            bool allQuantitiesValid = true; 
            string prodQtyColumnXPath = "//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]/div[10]"; // XPath pour la colonne Prod. Qty

            try
            {
                var element = WaitForElementIsVisible(By.XPath(prodQtyColumnXPath));

                if (element != null)
                {
                    string value = element.GetAttribute("value")?.Trim();

                    if (!string.IsNullOrEmpty(value))
                    {
                        // Essayer de convertir le texte en entier
                        if (int.TryParse(value, out int quantity))
                        {
                            // Vérifier si la quantité est négative
                            if (quantity < 0)
                            {
                                allQuantitiesValid = false; // Une quantité est négative
                            }
                        }
                        else
                        {
                            allQuantitiesValid = false; 
                        }
                    }
                }
                else
                {
                    allQuantitiesValid = false;
                }
            }
            catch (NoSuchElementException)
            {
                allQuantitiesValid = false;
            }

            // Retourner vrai si toutes les quantités sont valides (>= 0 ou valeur vide)
            return allQuantitiesValid;
        }


        public SupplyOrderItem SelectFirstItem()
        {
            _firstSONumber = WaitForElementIsVisible(By.XPath(FIRST_SO_NUMBER));
            _firstSONumber.Click();
            WaitForLoad(); 

            return new SupplyOrderItem(_webDriver, _testContext);
        }
        public void SelectFirstItemSO()
        {
            _firstSOItem = WaitForElementIsVisible(By.XPath(FIRST_SO_ITEM));
            _firstSOItem.Click();
            WaitForLoad();
        }

        public void DeleteFirstSupplyOrder()
        {
            WaitPageLoading();

            if (IsDev())
            {
                _deleteFirstSO = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody/tr[1]/td[8]/a"));
            }
            else
            {
                WaitForLoad();
                _deleteFirstSO = WaitForElementIsVisible(By.XPath(DELETE_FIRST_SO));
            }
            _deleteFirstSO.Click();
            WaitForLoad();

            _confirmDelete = WaitForElementIsVisible(By.Id(CONFIRM_DELETE));
            _confirmDelete.Click();
            WaitForLoad();
        }

        public PrintReportPage PrintResults(bool versionPrint, string downloadPath = null)
        {
            ShowExtendedMenu();
            _printSupplyOrders = WaitForElementIsVisible(By.Id(PRINT_RESULTS));
            _printSupplyOrders.Click();
            WaitForLoad();

            if (isElementVisible(By.XPath("//h4[contains(text(), 'Print')]"))) //new modal for include prices on report
            {
                _modalPrintButton = WaitForElementIsVisible(By.Id(MODAL_PRINT_BUTTON));
                _modalPrintButton.Click();
                WaitForLoad();
            }

            if (versionPrint)
            {
                var reportPage = PrintSupplyOrderNewVersion();
                var isGenerated = reportPage.IsReportGenerated();
                reportPage.Close();

                Assert.IsTrue(isGenerated, "Le document PDF n'a pas pu être généré par l'application.");
            }
            else
            {
                WaitForDownload();
                Close();

                // On récupère les fichiers du répertoire de téléchargement
                DirectoryInfo taskDirectory = new DirectoryInfo(downloadPath);
                FileInfo[] taskFiles = taskDirectory.GetFiles();

                // On exporte les résultats sous la forme d'un fichier Excel (dont on récupère le nom)
                // Export du fichier au format Excel
                var correctDownloadedFile = GetExportPdfFile(taskFiles);
                Assert.IsNotNull(correctDownloadedFile, MessageErreur.FICHIER_NON_TROUVE);
            }
            return new PrintReportPage(_webDriver, _testContext);
        }

        private PrintReportPage PrintSupplyOrderNewVersion()
        {
            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        public FileInfo GetExportPdfFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsPdfFileCorrect(file.Name))
                {
                    correctDownloadFiles.Add(file);
                }
            }
            Assert.IsTrue(correctDownloadFiles.Count > 0);

            var time = correctDownloadFiles[0].LastWriteTimeUtc;
            var correctFile = correctDownloadFiles[0];

            correctDownloadFiles.ForEach(file =>
            {
                if (time < file.LastWriteTimeUtc)
                {
                    time = file.LastWriteTimeUtc;
                    correctFile = file;
                }
            });

            return correctFile;
        }

        public bool IsPdfFileCorrect(string filePath)
        {
            // SupplyOrders_20200417_030319.pdf
            string mois = "(?:0[1-9]|1[0-2])";  // mois
            string annee = "\\d{4}";            // annee YYYY
            string jour = "[0-3]\\d";           // jour
            string nombre = "\\d{6}";           // nombre


            Regex r = new Regex("^SupplyOrders_" + annee + mois + jour + "_" + nombre + ".pdf$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public void Export(bool versionPrint)
        {
            ShowExtendedMenu();
            _export = WaitForElementIsVisible(By.Id(EXPORT));
            _export.Click();

            if(versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            }

            WaitForDownload();
            Close();
        }

        public string GetMessageExport()
        {
            ShowExtendedMenu();
            _export = WaitForElementIsVisible(By.Id(EXPORT));
            _export.Click();

            if (isElementVisible(By.XPath("//*[@id=\"dataAlertModal\"]/div/div/div[2]/p")))
            {
                var messagetext = WaitForElementIsVisible(By.XPath("//*[@id=\"dataAlertModal\"]/div/div/div[2]/p"));
                return messagetext.Text;
            }
            else
            {
                return "";
            }
        }

        public FileInfo GetExportExcelFile(FileInfo[] taskFiles)
        {
            WaitPageLoading();
            WaitLoading();
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsExcelFileCorrect(file.Name))
                {
                    correctDownloadFiles.Add(file);
                }
            }
            if (correctDownloadFiles.Count <= 0)
            {
                throw new Exception("L'application n'a pas trouvé le fichier recherché.");
            }

            var time = correctDownloadFiles[0].LastWriteTimeUtc;
            var correctFile = correctDownloadFiles[0];

            correctDownloadFiles.ForEach(file =>
            {
                if (time < file.LastWriteTimeUtc)
                {
                    time = file.LastWriteTimeUtc;
                    correctFile = file;
                }
            });

            return correctFile;
        }

        public bool IsExcelFileCorrect(string filePath)
        {
            // "supplyorders 2019-10-16 14-18-47.xlsx"

            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour
            string heure = "(?:0[0-9]|1[0-9]|2[0-3])"; // heure
            string minutes = "[0-5]\\d";               // minutes
            string secondes = "[0-5]\\d";              // secondes

            Regex r = new Regex("^supplyorders" + space + annee + "-" + mois + "-" + jour + space + heure + "-" + minutes + "-" + secondes + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }
        public SupplyOrderItem ClickFirstSupplyOrder()
        {
            WaitPageLoading();
            var firstsupplyorder = WaitForElementIsVisible(By.XPath(FIRST_SUPPLY_ORDER));
            firstsupplyorder.Click();
            WaitForLoad();
            return new SupplyOrderItem(_webDriver, _testContext);
        }

        public SupplyOrderItem ClickFirstSupplyOrderAfter()
        {
            WaitPageLoading();
            var firstsupplyorder = WaitForElementIsVisible(By.XPath(FIRST_SUPPLY_ORDER_NOTSUPPLIED));
            firstsupplyorder.Click();
            WaitForLoad();
            return new SupplyOrderItem(_webDriver, _testContext);
        }

        public void ClickShowItemsNotSupplied()
        {
             var checkboxElement = WaitForElementExists(By.Id(SHOWITEMSNOTSUPPLIED));

             bool isChecked = checkboxElement.Selected;

             if (!isChecked)
            {
                checkboxElement.Click();
                WaitForLoad();
            }
        }

        public string GetSearchValue()
        {
            var searchValue = WaitForElementIsVisible(By.Id(SEARCH_VALUE));
            return searchValue.Text;
        }

        public void DownloadReport()
        {
            var DownloadFile = WaitForElementExists(By.XPath(DOWNLOAD_FILE));
            DownloadFile.Click();
        }
        public IEnumerable<string> GetSupplyOrdersNumbersList()
        {
            var listnumbers = _webDriver.FindElements(By.XPath(SUPPLY_ORDERS_NUMBERS));
            var list = new List<string>();
            foreach(var number in listnumbers)
            {
                list.Add(number.Text);
            }
            return list;
        }
        public bool VerifyPdf(string path,string number)
        {
            PdfDocument document = PdfDocument.Open(path);
            List<string> mots = new List<string>();

            foreach (var p in document.GetPages())
            {
                foreach (var mot in p.GetWords())
                {
                    mots.Add(mot.Text);
                }
            }
            foreach(var mot in mots)
            {
                if(mot == number)
                {
                    return true;
                }
            }
           return false;
        }

        public string NombreSupplyOrdersS()
        {
            var nbsupplyorders = WaitForElementExists(By.XPath("//*[@id=\"div-body\"]/div/div[2]/div[1]/h1/span"));
            return nbsupplyorders.Text;
        }

        public int ConvertStringToInt(string value)
        {

            if (int.TryParse(value, out int validResult))
            {
                return validResult;
            }
            return 0;
        }

        public int GetTotalRowsForPagination()
        {
            var table = _webDriver.FindElement(By.XPath("//*[@id=\"list-item-with-action\"]/table"));

            var allRows = table.FindElements(By.TagName("tr"));

            var totalRows = allRows.Count() - 1;
            return totalRows;
        }

        public void ChooseDeliveryLocationOption(string deliveryLocOption)
        {
            ComboBoxSelectById(new ComboBoxOptions(SO_DELIVERY_LOCATION_SELECT, deliveryLocOption));
        }
        public void ClearDownloads()
        {
            int compteur = 1;
            bool isVisible = false;

            while (compteur <= 5 && !isVisible)
            {
                var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(5));

                ClickPrintButton();
                try
                {
                    var clearDownloadButton = wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.CssSelector("[onclick=\"clearPrintList()\"]")));
                    clearDownloadButton.Click();
                    WaitForLoad();
                    isVisible = true;

                    ClosePrintButton();
                    // Ophélie : Temps de fermeture de la fenêtre
                    //_webDriver.Navigate().Refresh();
                }
                catch
                {
                    ClickPrintButton();
                    //David : Double clic trop rapide si on n'attend pas
                    Thread.Sleep(500);
                    compteur++;
                }
            }

            if (!isVisible)
            {
                Console.WriteLine($"[{DateTime.UtcNow.ToShortDateString()} {DateTime.UtcNow.ToLongTimeString()}] [ClearDownloads] Clear button not visible");
                throw new Exception("La fonction 'Clear' du Print n'est pas accessible.");
            }
        }
        public bool IsValidated()
        {
            var modalElement = _webDriver.FindElement(By.Id("purchasing-supplyOrder-isvalid-1"));
            return modalElement.Displayed;
        }
        public SupplyOrderItem SelectFirstNonValidatedItem()
        {
            // Récupérer tous les éléments des items de la Supply Order
            var items = _webDriver.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/table/tbody"));

            foreach (var item in items)
            {
                try
                {
                    // Vérifier si l'icône de validation est présente
                    var validationIcon = item.FindElement(By.CssSelector("color: limegreen; padding: 5px;"));

                    // Si l'icône est affichée, cet item est validé, on passe au suivant
                    if (!validationIcon.Displayed)
                    {
                        // Si l'icône n'est pas affichée, on clique sur cet item
                        item.Click();
                        return new SupplyOrderItem(_webDriver, _testContext);
                    }
                }
                catch (NoSuchElementException)
                {
                    // Si l'icône de validation n'est pas trouvée, cela signifie que l'item n'est pas validé
                    item.Click();
                    return new SupplyOrderItem(_webDriver, _testContext);
                }
            }

            // Si aucun item non validé n'est trouvé, lever une exception
            throw new NoSuchElementException("Aucun item non validé trouvé.");
        }
        public string GetItemName()
        {
            var itemNameElement = _webDriver.FindElement(By.XPath("//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[1]/div[3]/span"));
            return itemNameElement.Text.Trim();
        }

        public string GetItemQuantityAndUnit()
        {
            var itemQtyUnitElement = _webDriver.FindElement(By.XPath(QTY_UNIT));
            return itemQtyUnitElement.Text.Trim();
        }

        public string GetItemQuantityProd()
        {
            WaitPageLoading();
            WaitForLoad();
            WaitForLoad();
            var itemQtyItemElement = _webDriver.FindElement(By.XPath(QTY_ITEM));
            return itemQtyItemElement.Text.Trim();
        }

        
        public (string, string) SplitQuantityAndUnit(string combinedQtyUnit)
        {
            // Identifier l'index de la parenthèse ouvrante
            int parenIndex = combinedQtyUnit.IndexOf('(');

            // Si une parenthèse est trouvée
            if (parenIndex > 0)
            {
                // Extraire la portion avant la parenthèse (par exemple, "10 KG")
                string qtyUnitPart = combinedQtyUnit.Substring(0, parenIndex).Trim();

                // Séparer cette portion par espace pour obtenir la quantité et l'unité
                var parts = qtyUnitPart.Split(' ');

                if (parts.Length == 2)
                {
                    string quantity = parts[0].Trim();
                    string unit = parts[1].Trim();
                    return (quantity, unit);
                }
                else
                {
                    throw new InvalidOperationException("Format quantité/unité invalide");
                }
            }
            else
            {
                // Si aucune parenthèse n'est trouvée, traiter la chaîne entière
                var parts = combinedQtyUnit.Trim().Split(' ');

                if (parts.Length == 2)
                {
                    string quantity = parts[0].Trim();
                    string unit = parts[1].Trim();
                    return (quantity, unit);
                }
                else
                {
                    throw new InvalidOperationException("Format quantité/unité invalide");
                }
            }
        }

        public bool IsSupplyOrderCreatedSuccessfully()
        {
            try
            {
                // XPath mis à jour pour cibler la table ou le conteneur de la liste
                var orderContainer = _webDriver.FindElement(By.XPath("//*[@id=\"list-item-with-action\"]/div"));

                // Suppose que les lignes sont des <div> avec une certaine classe
                var items = orderContainer.FindElements(By.XPath("//*[@id=\"list-item-with-action\"]/div/div[2]/div[2]/div/div/form/div[2]"));

                // Vérifie s'il y a au moins une ligne/item
                return items.Count > 0;
            }
            catch (NoSuchElementException)
            {
                // Table ou éléments de SO non trouvés
                return false;
            }
        }
      
        public bool CheckTotalCalculated(string totalCalculated , string dispatchQty , string packingPrice)
        {
            var qty= dispatchQty == string.Empty? "0": dispatchQty;
            var price = packingPrice == string.Empty ? "0" : packingPrice;
            
           return Convert.ToDouble(totalCalculated) == Convert.ToDouble(qty) * Convert.ToDouble(price);
        }
    }
}
