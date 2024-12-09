using Microsoft.VisualBasic.FileIO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Parameters.Production;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Production.Setup
{
    public class SetupPage : PageBase
    {

        public SetupPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //__________________________________ Constantes _____________________________________

        // General

        // Views

        // Tableau

        // Onglets
        private const string DELIVERY_ROUND_TAB = "hrefTabContentItemContainer";

        // Filtres
        private const string RESET_FILTER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[1]/a";
        private const string SEARCH_DELIVERY_ROUND_NAME_INPUT = "SearchPattern";
        private const string SETUP_FILTER_DATE = "StartDate";
        private const string SITES = "SiteId";
        private const string SEARCH_RECIPE_NAME_INPUT = "SearchRecipeName";
        private const string CUSTOMER_TYPES_FILTER = "SelectedCustomerTypesIds_ms";
        private const string CUSTOMER_TYPES_FILTER_UNCHECK_ALL = "/html/body/div[10]/div/ul/li[2]/a/span[2]";
        private const string CUSTOMER_TYPES_FILTER_SEARCH = "/html/body/div[10]/div/div/label/input";
        private const string CUSTOMERS_FILTER = "SelectedCustomersIds_ms";
        private const string CUSTOMERS_FILTER_UNCHECK_ALL = "/html/body/div[11]/div/ul/li[2]/a/span[2]";
        private const string CUSTOMERS_FILTER_SEARCH = "/html/body/div[11]/div/div/label/input";

        private const string DELIVERIES_FILTER = "SelectedDeliveriesIds_ms";

        private const string MEALTYPES_FILTER = "SelectedMealTypesIds_ms";

        private const string WORKSHOPS_FILTER = "SelectedWorkshopsIds_ms";

        private const string NUMBER_OF_SETUP = "//*[@id=\"tabContentItemContainer\"]/div[1]/h1/span";

        // Print
        private const string BTN_PRINT_DELIVERYROUND = "printSelection_10";
        private const string BTN_PRINT_DELIVERYROUNDBYRECIPES = "printSelection_12";
        private const string BTN_PRINT_DELIVERYNOTEBYRECIPES = "printSelection_15";
        private const string BTN_PRINT_DELIVERYNOTEBYSERVICES = "printSelection_16";
        private const string BTN_PRINT_DELIVERYNOTEVALORIZED = "printSelection_14";
        private const string BTN_PRINT_FOODPACKREPORT = "printSelection_17";
        private const string BTN_PRINT_FOODPACKGROUPBYDELIVERY = "printSelection_18";
        private const string BTN_PRINT_EXPORT = "ExportSelection";
        private const string BTN_PRINT_PARAMETERS = "gotoParameters";

        private const string VALIDATE_PRINT = "validatePrint";
        private const string EXPORT_VALIDATION = "validateExport";
        private const string DELIVERY = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td";

        //__________________________________ Variables _____________________________________

        // General

        // Views

        // Tableau

        // Onglets

        [FindsBy(How = How.Id, Using = DELIVERY_ROUND_TAB)]
        private IWebElement _deliveryRoundTab;

        //Filters
        [FindsBy(How = How.XPath, Using = RESET_FILTER)]
        private IWebElement _resetFilter;

        [FindsBy(How = How.Id, Using = SEARCH_DELIVERY_ROUND_NAME_INPUT)]
        private IWebElement _searchDeliveryRoundNameInput;

        [FindsBy(How = How.Id, Using = SETUP_FILTER_DATE)]
        private IWebElement _setupFilterDate;

        [FindsBy(How = How.Id, Using = SITES)]
        private IWebElement _sites;

        [FindsBy(How = How.Id, Using = SEARCH_RECIPE_NAME_INPUT)]
        private IWebElement _searchRecipeNameInput;

        [FindsBy(How = How.Id, Using = CUSTOMER_TYPES_FILTER)]
        private IWebElement _customerTypesFilter;

        [FindsBy(How = How.XPath, Using = CUSTOMER_TYPES_FILTER_UNCHECK_ALL)]
        private IWebElement _customerTypesFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = CUSTOMER_TYPES_FILTER_SEARCH)]
        private IWebElement _customerTypesFilterSearch;

        [FindsBy(How = How.Id, Using = CUSTOMERS_FILTER)]
        private IWebElement _customersFilter;

        [FindsBy(How = How.XPath, Using = CUSTOMERS_FILTER_UNCHECK_ALL)]
        private IWebElement _customersFilterUncheckAll;

        [FindsBy(How = How.XPath, Using = CUSTOMERS_FILTER_SEARCH)]
        private IWebElement _customersFilterSearch;

        [FindsBy(How = How.Id, Using = DELIVERIES_FILTER)]
        private IWebElement _deliveriesFilter;

        [FindsBy(How = How.Id, Using = MEALTYPES_FILTER)]
        private IWebElement _mealtypesFilter;

        [FindsBy(How = How.Id, Using = WORKSHOPS_FILTER)]
        private IWebElement _workshopsFilter;

        [FindsBy(How = How.XPath, Using = NUMBER_OF_SETUP)]
        private IWebElement _numberOfSetup;

        // Prints
        [FindsBy(How = How.Id, Using = BTN_PRINT_DELIVERYROUND)]
        private IWebElement _printDeliveryround;

        [FindsBy(How = How.Id, Using = BTN_PRINT_DELIVERYROUNDBYRECIPES)]
        private IWebElement _printDeliveryroundByRecipes;

        [FindsBy(How = How.Id, Using = BTN_PRINT_DELIVERYNOTEBYRECIPES)]
        private IWebElement _printDeliveryNoteByRecipes;

        [FindsBy(How = How.Id, Using = BTN_PRINT_DELIVERYNOTEBYSERVICES)]
        private IWebElement _printDeliveryNoteByServices;

        [FindsBy(How = How.Id, Using = BTN_PRINT_DELIVERYNOTEVALORIZED)]
        private IWebElement _printDeliveryNoteValorized;

        [FindsBy(How = How.Id, Using = BTN_PRINT_FOODPACKREPORT)]
        private IWebElement _printFoodPackReport;

        [FindsBy(How = How.Id, Using = BTN_PRINT_FOODPACKGROUPBYDELIVERY)]
        private IWebElement _printFoodPackGroupByDelivery;

        [FindsBy(How = How.Id, Using = VALIDATE_PRINT)]
        private IWebElement _validatePrint;

        // Export
        [FindsBy(How = How.Id, Using = BTN_PRINT_EXPORT)]
        private IWebElement _export;

        [FindsBy(How = How.Id, Using = EXPORT_VALIDATION)]
        private IWebElement _exportValidation;

        // Go to parameters
        [FindsBy(How = How.Id, Using = BTN_PRINT_PARAMETERS)]
        private IWebElement _goToSettings;

        //__________________________________ Methods _____________________________________
        // ONGLETS
        public SetupDeliveryRoundTabPage GoToSetupDeliveryRoundTab()
        {
            WaitForLoad();
            _deliveryRoundTab = WaitForElementIsVisible(By.Id(DELIVERY_ROUND_TAB));
            _deliveryRoundTab.Click();
            WaitForLoad();
            return new SetupDeliveryRoundTabPage(_webDriver, _testContext);
        }

        // FILTRES

        public enum FilterType
        {
            SearchDeliveryRoundName,
            StartDate,
            Sites,
            RecipeName,
            CustomerTypes,
            Customers,
            Deliveries,
            MealTypes,
            Workshops
        }
        public void ResetFilters()
        {
            _resetFilter = WaitForElementIsVisible(By.XPath(RESET_FILTER));
            _resetFilter.Click();
            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                // pas de date de fin
            }
        }
        public void Filter(FilterType filterType, object value)
        {
            Actions action = new Actions(_webDriver);

            switch (filterType)
            {
                case FilterType.SearchDeliveryRoundName:
                    _searchDeliveryRoundNameInput = WaitForElementIsVisible(By.Id(SEARCH_DELIVERY_ROUND_NAME_INPUT));
                    _searchDeliveryRoundNameInput.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;
                case FilterType.StartDate:
                    _setupFilterDate = WaitForElementIsVisible(By.Id(SETUP_FILTER_DATE));
                    _setupFilterDate.SetValue(ControlType.DateTime, value);
                    _setupFilterDate.SendKeys(Keys.Tab);
                    WaitForLoad();
                    break;
                case FilterType.Sites:
                    _sites = WaitForElementIsVisible(By.Id(SITES));
                    _sites.SetValue(ControlType.DropDownList, value);
                    break;
                case FilterType.RecipeName:
                    _searchRecipeNameInput = WaitForElementIsVisible(By.Id(SEARCH_RECIPE_NAME_INPUT));
                    _searchRecipeNameInput.SetValue(ControlType.TextBox, value);
                    WaitForLoad();
                    break;
                case FilterType.CustomerTypes:
                    _customerTypesFilter = WaitForElementIsVisible(By.Id(CUSTOMER_TYPES_FILTER));
                    _customerTypesFilter.Click();

                    _customerTypesFilterUncheckAll = WaitForElementIsVisible(By.XPath(CUSTOMER_TYPES_FILTER_UNCHECK_ALL));
                    _customerTypesFilterUncheckAll.Click();

                    _customerTypesFilterSearch = WaitForElementIsVisible(By.XPath(CUSTOMER_TYPES_FILTER_SEARCH));
                    _customerTypesFilterSearch.SetValue(ControlType.TextBox, value);

                    var _customerTypeToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    _customerTypeToCheck.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;
                case FilterType.Customers:
                    _customerTypesFilter = WaitForElementIsVisible(By.Id(CUSTOMERS_FILTER));
                    _customerTypesFilter.Click();

                    _customerTypesFilterUncheckAll = WaitForElementIsVisible(By.XPath(CUSTOMERS_FILTER_UNCHECK_ALL));
                    _customerTypesFilterUncheckAll.Click();

                    _customerTypesFilterSearch = WaitForElementIsVisible(By.XPath(CUSTOMERS_FILTER_SEARCH));
                    _customerTypesFilterSearch.SetValue(ControlType.TextBox, value);

                    var _customerToCheck = WaitForElementIsVisible(By.XPath("//span[text()='" + value + "']"));
                    _customerToCheck.SetValue(ControlType.CheckBox, true);
                    WaitForLoad();
                    break;
                case FilterType.Deliveries:
                    ComboBoxSelectById(new ComboBoxOptions(DELIVERIES_FILTER, (string)value));
                    break;
                case FilterType.MealTypes:
                    ComboBoxSelectById(new ComboBoxOptions(MEALTYPES_FILTER, (string)value));
                    break;
                case FilterType.Workshops:
                    new Actions(_webDriver).MoveToElement(WaitForElementExists(By.Id(WORKSHOPS_FILTER))).Perform();
                    ComboBoxSelectById(new ComboBoxOptions(WORKSHOPS_FILTER, (string)value));
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(FilterType), filterType, null);

            }

            WaitPageLoading();

            try
            {
                WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
                wait.Until(ExpectedConditions.ElementIsVisible(By.Id("hrefTabContentItemContainer")));
            }
            catch
            {
                throw new Exception($"Le tableau n'affiche pas de résultats dans les 60sec");
            }
        }
        public bool IsResultsDisplayed()
        {
            Thread.Sleep(2000);//wait table
            _numberOfSetup = WaitForElementIsVisible(By.XPath(NUMBER_OF_SETUP));
            if (_numberOfSetup.Text.Equals("0"))
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        public enum PrintType
        {
            DeliveryRound,
            DeliveryRoundByRecipes,
            DeliveryNoteByRecipes,
            DeliveryNoteByServices,
            DeliveryNoteValorized,
            FoodPackReport,
            FoodPackGroupByDelivery,
            Export,
            Parameters
        }
        public PrintReportPage Print(PrintType printType, bool printValue)
        {
            ShowExtendedMenu();

            switch (printType)
            {
                case PrintType.DeliveryRound:
                    _printDeliveryround = WaitForElementIsVisible(By.Id(BTN_PRINT_DELIVERYROUND));
                    _printDeliveryround.Click();
                    WaitForLoad();
                    break;

                case PrintType.DeliveryRoundByRecipes:
                    _printDeliveryroundByRecipes = WaitForElementIsVisible(By.Id(BTN_PRINT_DELIVERYROUNDBYRECIPES));
                    _printDeliveryroundByRecipes.Click();
                    WaitForLoad();
                    break;

                case PrintType.DeliveryNoteByRecipes:
                    _printDeliveryNoteByRecipes = WaitForElementIsVisible(By.Id(BTN_PRINT_DELIVERYNOTEBYRECIPES));
                    _printDeliveryNoteByRecipes.Click();
                    WaitForLoad();
                    break;

                case PrintType.DeliveryNoteByServices:
                    _printDeliveryNoteByServices = WaitForElementIsVisible(By.Id(BTN_PRINT_DELIVERYNOTEBYSERVICES));
                    _printDeliveryNoteByServices.Click();
                    WaitForLoad();
                    break;

                case PrintType.DeliveryNoteValorized:
                    _printDeliveryNoteValorized = WaitForElementIsVisible(By.Id(BTN_PRINT_DELIVERYNOTEVALORIZED));
                    _printDeliveryNoteValorized.Click();
                    WaitForLoad();
                    break;

                case PrintType.FoodPackReport:
                    _printFoodPackReport = WaitForElementIsVisible(By.Id(BTN_PRINT_FOODPACKREPORT));
                    _printFoodPackReport.Click();
                    WaitForLoad();
                    break;

                case PrintType.FoodPackGroupByDelivery:
                    _printFoodPackGroupByDelivery = WaitForElementIsVisible(By.Id(BTN_PRINT_FOODPACKGROUPBYDELIVERY));
                    _printFoodPackGroupByDelivery.Click();
                    WaitForLoad();
                    break;
            }

            _validatePrint = WaitForElementIsVisible(By.Id(VALIDATE_PRINT));
            _validatePrint.Click();
            WaitForLoad();

            if (printValue)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();
                WaitForLoad();
            }

            //Results are opened in a new tab, switch the driver to the newly created one
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        public void ExportCSV(bool versionPrint)
        {
            ShowExtendedMenu();

            _export = WaitForElementIsVisible(By.Id(BTN_PRINT_EXPORT));
            _export.Click();
            WaitForLoad();

            _exportValidation = WaitForElementIsVisible(By.Id(EXPORT_VALIDATION));
            _exportValidation.Click();
            WaitForLoad();

            if (versionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-alt']"));
            }

            WaitForDownload();
            Close();
        }

        public FileInfo GetCsvFile(FileInfo[] taskFiles)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();

            foreach (var file in taskFiles)
            {
                if (IsCSVFileNameCorrect(file.Name))
                {
                    correctDownloadFiles.Add(file);
                }
            }

            if (correctDownloadFiles.Count <= 0)
            {
                throw new Exception(MessageErreur.FICHIER_NON_TROUVE);
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

        public bool IsCSVFileNameCorrect(string filePath)
        {
            // "Setup 2021-06-29 15-23-12.csv"
            string today = DateUtils.Now.ToString("yyyy-MM-dd");

            Regex r = new Regex("Setup " + today + "(.*).csv$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Match m = r.Match(filePath);

            return m.Success;
        }

        public string ReadCSVFile(string filepath, string value)
        {
            List<string> headerList = new List<string>();
            List<string> valuesList = new List<string>();

            using (var reader = new StreamReader(filepath))
            {
                var line = reader.ReadLine().Split(';');
                foreach (var item in line)
                {
                    headerList.Add(item.Replace("\"", ""));
                }
                line = reader.ReadLine().Split(';');
                foreach (var item in line)
                {
                    valuesList.Add(item.Replace("\"", ""));
                }

                int index = headerList.IndexOf(value);
                var valueCorres = valuesList[index];
                return valueCorres;
            }
        }

        public List<List<string>> GetCSVData(string filePath)
        {
            List<List<string>> csvData = new List<List<string>>();

            // Get csv data all lines
            using (TextFieldParser parser = new TextFieldParser(filePath))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(";");
                while (!parser.EndOfData)
                {
                    //Process row
                    var values = new List<string>();

                    var readFields = parser.ReadFields();
                    if (readFields != null)
                        values.AddRange(readFields);
                    csvData.Add(values);
                }
            }
            return csvData;
        }

        public string GetNbLabel(List<List<string>> csvData, string recipeName, string nbPAX, string foodpackName)
        {
            int nbLabelIndex = csvData[0].IndexOf("Nb label");
            int recipeNameIndex = csvData[0].IndexOf("Recipe name");
            int nbPAXIndex = csvData[0].IndexOf("Nb PAX /Packaging");
            int foodPackNameIndex = csvData[0].IndexOf("Food pack name");

            var line = csvData
                .Where(list => list[recipeNameIndex].Contains(recipeName))
                .Where(list => list[nbPAXIndex].Contains(nbPAX))
                .Where(list => list[foodPackNameIndex].Equals(foodpackName))
                .ToList();
            return line[0][nbLabelIndex];
        }

        public string GetSiteName(List<List<string>> csvData, string recipeName, string menuName, string foodpackName)
        {
            int siteNameIndex = csvData[0].IndexOf("Site name");
            int recipeNameIndex = csvData[0].IndexOf("Recipe name");
            int menuNameIndex = csvData[0].IndexOf("Menu name");
            int foodPackNameIndex = csvData[0].IndexOf("Food pack name");
            try
            {
                var line = csvData
                    .Where(list => list[recipeNameIndex].Contains(recipeName))
                    .Where(list => list[menuNameIndex].Contains(menuName))
                    .Where(list => list[foodPackNameIndex].Equals(foodpackName))
                    .ToList();
                return line[0][siteNameIndex];
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Assert.Fail(recipeNameIndex + ":" + recipeName + " " + menuNameIndex + ":" + menuName + " " + foodPackNameIndex + ":" + foodpackName);
                return "";
            }
        }

        public List<string> GetSiteAddress(List<List<string>> csvData, string siteName)
        {
            var siteAddressList = new List<string>();
            int siteAddressIndex = csvData[0].IndexOf("Site address");
            int siteNameIndex = csvData[0].IndexOf("Site name");

            var excelLines = csvData
                .Where(list => list[siteNameIndex].Contains(siteName))
                .ToList();
            foreach (var line in excelLines)
            {
                siteAddressList.Add(line[siteAddressIndex]);

            }
            return siteAddressList.Distinct().ToList();
        }

        public List<string> GetSiteSanitaryAgreement(List<List<string>> csvData, string siteName)
        {
            var siteSanitaryAgreementList = new List<string>();
            int siteSanitaryAgreementIndex = csvData[0].IndexOf("Sanitary agreement");
            int siteNameIndex = csvData[0].IndexOf("Site name");

            var excelLines = csvData
                .Where(list => list[siteNameIndex].Contains(siteName))
                .ToList();
            foreach (var line in excelLines)
            {
                siteSanitaryAgreementList.Add(line[siteSanitaryAgreementIndex]);

            }
            return siteSanitaryAgreementList.Distinct().ToList();
        }

        public List<string> GetExportResultsByRecipe(List<List<string>> csvData, string columnName, string recipeName)
        {
            var resultList = new List<string>();
            int columnNameIndex = csvData[0].IndexOf(columnName);
            int recipeNameIndex = csvData[0].IndexOf("Recipe name");

            var excelLines = csvData
                .Where(list => list[recipeNameIndex].Contains(recipeName))
                .ToList();
            foreach (var line in excelLines)
            {
                resultList.Add(line[columnNameIndex]);
            }
            return resultList.Distinct().ToList();
        }
        public List<string> GetExportResults(List<List<string>> csvData, string columnName)
        {
            var resultList = new List<string>();
            int columnNameIndex = csvData[0].IndexOf(columnName);

            var excelLines = csvData
                .ToList();
            foreach (var line in excelLines)
            {
                if (line[columnNameIndex] != columnName)
                {
                    resultList.Add(line[columnNameIndex]);
                }
            }
            return resultList.Distinct().ToList();
        }
        public bool IsMenuNameInExcel(List<List<string>> csvData, string partialMenuName)
        {
            if (csvData == null || csvData.Count == 0)
            {
                throw new ArgumentException("The CSV data is empty or null.");
            }
            int menuNameIndex = csvData[0].IndexOf("Menu name");
            if (menuNameIndex == -1)
            {
                throw new ArgumentException("The column 'Menu name' was not found in the CSV data.");
            }
            var matchingRows = csvData
                .Skip(1) // Skip the header row
                .Where(row => row.Count > menuNameIndex &&
                              row[menuNameIndex].IndexOf(partialMenuName, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();
            return matchingRows.Count > 0;
        }


        public List<string> GetRecipesNamesList(List<List<string>> csvData, string menuName)
        {
            var recipesList = new List<string>();
            int recipeNameIndex = csvData[0].IndexOf("Recipe name");
            int menuNameIndex = csvData[0].IndexOf("Menu name");

            var excelLines = csvData
                .Where(list => list[menuNameIndex].Contains(menuName))
                .ToList();
            foreach (var line in excelLines)
            {
                recipesList.Add(line[recipeNameIndex]);

            }
            return recipesList.Distinct().ToList();
        }

        public List<string> GetCommercialRecipesNamesList(List<List<string>> csvData, string menuName)
        {
            var commercialRecipesList = new List<string>();
            int commercialRecipeNameIndex = csvData[0].IndexOf("Commercial name");
            int menuNameIndex = csvData[0].IndexOf("Menu name");

            var excelLines = csvData
                .Where(list => list[menuNameIndex].Contains(menuName))
                .ToList();
            foreach (var line in excelLines)
            {
                commercialRecipesList.Add(line[commercialRecipeNameIndex]);

            }
            return commercialRecipesList.Distinct().ToList();
        }
        public string GetGuestType(List<List<string>> csvData, string menuName)
        {
            int guestTypeIndex = csvData[0].IndexOf("Guest type");
            int menuNameIndex = csvData[0].IndexOf("Menu name");

            var line = csvData
                .Where(list => list[menuNameIndex].Contains(menuName))
                .ToList();
            if (line.Count == 0)
            {
                return "";
            }
            else
            {
                return line[0][guestTypeIndex];
            }
        }
        public string GetMealType(List<List<string>> csvData, string menuName)
        {
            int mealTypeIndex = csvData[0].IndexOf("Meal type");
            int menuNameIndex = csvData[0].IndexOf("Menu name");

            var line = csvData
                .Where(list => list[menuNameIndex].Contains(menuName))
                .ToList();

            if (line.Count == 0)
            {
                return "";
            }
            else
            {
                return line[0][mealTypeIndex];
            }
        }

        public string GetNbPAXPerPackaging(List<List<string>> csvData, string menuName, string foodpackName)
        {
            int nbPAXPerPackagingIndex = csvData[0].IndexOf("Nb PAX /Packaging");
            int menuNameIndex = csvData[0].IndexOf("Menu name");
            int foodPackNameIndex = csvData[0].IndexOf("Food pack name");

            var line = csvData
                .Where(list => list[menuNameIndex].Contains(menuName))
                .Where(list => list[foodPackNameIndex].Equals(foodpackName))
                .ToList();

            if (line.Count == 0)
            {
                return "";
            }
            else
            {
                return line[0][nbPAXPerPackagingIndex];
            }
        }

        public string GetFoodPackName(List<List<string>> csvData, string menuName, string nbPAXPerPackaging)
        {
            int foodPackNameIndex = csvData[0].IndexOf("Food pack name");
            int menuNameIndex = csvData[0].IndexOf("Menu name");
            int nbPAXPerPackagingIndex = csvData[0].IndexOf("Nb PAX /Packaging");

            var line = csvData
                .Where(list => list[menuNameIndex].Contains(menuName))
                .Where(list => list[nbPAXPerPackagingIndex].Equals(nbPAXPerPackaging))
                .ToList();

            if (line.Count == 0)
            {
                return "";
            }
            else
            {
                return line[0][foodPackNameIndex];
            }
        }

        public ParametersSetup GoToParameters_SetupSettingsFromSetupPage()
        {
            ShowExtendedMenu();

            _goToSettings = WaitForElementIsVisible(By.Id(BTN_PRINT_PARAMETERS));
            _goToSettings.Click();
            WaitPageLoading();

            return new ParametersSetup(_webDriver, _testContext);
        }
        public string GetCurrentUrl()
        {
            return _webDriver.Url;
        }


        public bool HasResults()
        {

            // Localiser l'élément correspondant à vos résultats (adapter l'XPath selon votre structure)
            var resultElement = _webDriver.FindElement(By.XPath("/html/body/div[2]/div/div[2]/div[1]/h1"));

            // Vérifier si le texte de l'élément est présent et non vide
            return !string.IsNullOrWhiteSpace(resultElement.Text);
        }
        public bool isFirstDeliveryRoundNameExist(string deliveryName)
        {
            return isElementExists(By.XPath($"//*[@id=\"list-item-with-action\"]/div[*]/div[1]/div/div[2]/table/tbody/tr/td[contains(text(),'{deliveryName}')]"));
        }
      public string GetDeliveryRoundName()
        
        {
            var element = WaitForElementIsVisible(By.XPath("//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[2]/table/tbody/tr/td"));
            return element.Text;
        }
        public bool IsPresent()
        {
            var _delivery = _webDriver.FindElements(By.XPath(DELIVERY)).Count;

            if (_delivery == 0)
                return false;

            return true;
        }
    }
}
