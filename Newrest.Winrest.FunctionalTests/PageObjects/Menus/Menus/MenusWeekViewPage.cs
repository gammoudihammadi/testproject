using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Vml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Menus.Recipes;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition.Primitives;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace Newrest.Winrest.FunctionalTests.PageObjects.Menus.Menus
{
    public class MenusWeekViewPage : PageBase
    {

        public MenusWeekViewPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        // ______________________________________ Constantes _____________________________________________

        // Général
        private const string MENU_TITLE = "//*[@id=\"div-body\"]/div[1]/div[1]/h1";
        private const string MENU_DATES = "//*[@id=\"div-body\"]/div[1]/div[2]/h2";
        private const string DAY_VIEW = "//*[@id=\"div-body\"]/div[1]/div[1]/div/a[4]";

        private const string EXPORT_DEV = "exportDataSheet";
        private const string EXPORT_PATCH = "//*[@id=\"div-body\"]/div[1]/div[1]/div/a[2]";

        private const string SATURDAY = "IsActiveOnSaturday";
        private const string SUNDAY = "IsActiveOnSunday";
        private const string CONFIRM_EXPORT = "print-export";
        private const string EXPORTMENUS = "exportBtn";


        
        // Add recipe
        private const string ADD_RECIPE = "RecipeName";
        private const string ACTIVE_DAY = "//*[@id=\"modalDailyCalendar\"]/div[*]/div[*]/div[*]/p[@data-real-date='{0}']";
        private const string CONFIRM_ADD = "//*[@id=\"modal-1\"]/div/div/div[2]/div/form/div[2]/button[2]";
        private const string NEXT_MONTH = "//*[@id=\"modalDailyCalendar\"]/div[4]/div[1]/button";
        private const string ACTIVE_SELECTED_DAY = "//*[@id=\"modalDailyCalendar\"]/div[*]/div[*]/div[*]/p[@data-real-date='{0}' and not(contains(@class, 'disabled'))]";

        // Tableau
        private const string LEFT_ARROW = "//*[@id=\"topTabWeekly\"]/div[2]/a[1]/span";
        private const string RIGHT_ARROW = "//*[@id=\"topTabWeekly\"]/div[2]/a[2]/span";
        private const string SEE_RECIPE_DETAIL = "//*[starts-with(@id,\"popover\")]/div[2]/div/div/div[4]/div/a";
        private const string METHOD = "MethodPopover";
        private const string METHOD_SELECTED = "//*[@id=\"MethodPopover\"]/option[@selected = 'selected']";

        private const string METHOD_SELECTED_2 = "/html/body/div[11]/div[2]/div/div/div[5]/div/select";
        private const string COEFFICIENT = "CoefficientPopover";
        private const string PRICE = "pricePopover";
        private const string SELECT_BTN = "//*[starts-with(@id,\"popover\")]/div[2]/div/div/button[1]";

        private const string MULTI_DELETE = "//*[@id=\"topTabWeekly\"]/div[1]/a[@title='Delete recipe']";
        private const string REMOVE_BTN = "//*[starts-with(@id,\"popover\")]/div[2]/div/div/button[2]";

        private const string MULTI_DUPLICATE = "//*[@id=\"topTabWeekly\"]/div[1]/a[@title='Duplicate selection']";
        private const string DUPLICATE_BTN = "//*[starts-with(@id,\"popover\")]/div[2]/div/div/button[3]";
        private const string FLECHE_DUPLICATE = "//*[@id=\"modal-1\"]/div[2]/div[1]/table/thead/tr[1]/th[3]";
        private const string DUPLICATE_DATE = "//*[@id=\"modal-1\"]/div[2]/div[1]/table/tbody/tr[*]/td[text() = '{0}']";
        private const string CONFIRM_DUPLICATE = "//*[@id=\"modal-1\"]/div[2]/div/form/div[2]/button[2]";

        private const string CLOSE = "//*[starts-with(@id,\"popover\")]/h3/span";
        private const string DATE_TABLEAU = "//*[@id=\"tMenuDaily\"]/tbody/tr[2]/td[text() = '{0}']";
        private const string RECIPE_ADDED = "//*[@id=\"tMenuDaily\"]/tbody/tr[3]/td[{0}]/ul/li[@data-recipe-name = '{1}']";

        private const string DAYS_LIST_WEEK_VIEW = "/html/body/div[3]/div[1]/div[3]/div/div[1]/table/tbody/tr[2]/td[*]";
        private const string DAYS_ELEMENT_LIST_WEEK_VIEW = "/html/body/div[3]/div[1]/div[3]/div/div[1]/table/tbody/tr[3]/td[*]";

        // ______________________________________ Variables ______________________________________________

        // Général
        [FindsBy(How = How.XPath, Using = MENU_DATES)]
        private IWebElement _menusDates;

        [FindsBy(How = How.XPath, Using = DAY_VIEW)]
        private IWebElement _dayView;

        [FindsBy(How = How.Id, Using = EXPORT_DEV)]
        private IWebElement _exportDev;
        
        [FindsBy(How = How.XPath, Using = EXPORT_PATCH)]
        private IWebElement _exportPatch;

        [FindsBy(How = How.XPath, Using = EXPORTMENUS)]
        private IWebElement _exportmenus;
        
        [FindsBy(How = How.Id, Using = SATURDAY)]
        private IWebElement _saturday;

        [FindsBy(How = How.Id, Using = SUNDAY)]
        private IWebElement _sunday;

        [FindsBy(How = How.Id, Using = CONFIRM_EXPORT)]
        private IWebElement _confirmExport;

        // Add recipe
        [FindsBy(How = How.Id, Using = ADD_RECIPE)]
        private IWebElement _addRecipe;

        [FindsBy(How = How.XPath, Using = CONFIRM_ADD)]
        private IWebElement _confirmAdd;

        // Tableau
        [FindsBy(How = How.XPath, Using = LEFT_ARROW)]
        private IWebElement _leftArrow;

        [FindsBy(How = How.XPath, Using = RIGHT_ARROW)]
        private IWebElement _rightArrow;

        [FindsBy(How = How.XPath, Using = REMOVE_BTN)]
        private IWebElement _removeRecipe;

        [FindsBy(How = How.XPath, Using = MULTI_DELETE)]
        private IWebElement _deleteMultiRecipe;

        [FindsBy(How = How.XPath, Using = DUPLICATE_BTN)]
        private IWebElement _duplicateRecipe;

        [FindsBy(How = How.XPath, Using = FLECHE_DUPLICATE)]
        private IWebElement _flecheDuplicate;

        [FindsBy(How = How.XPath, Using = MULTI_DUPLICATE)]
        private IWebElement _duplicateMultiRecipes;

        [FindsBy(How = How.XPath, Using = CONFIRM_DUPLICATE)]
        private IWebElement _confirmDuplicate;

        [FindsBy(How = How.XPath, Using = SELECT_BTN)]
        private IWebElement _selectRecipe;

        [FindsBy(How = How.XPath, Using = SEE_RECIPE_DETAIL)]
        private IWebElement _editRecipe;

        [FindsBy(How = How.XPath, Using = CLOSE)]
        private IWebElement _close;

        [FindsBy(How = How.Id, Using = METHOD)]
        private IWebElement _recipeMethod;

        [FindsBy(How = How.XPath, Using = METHOD_SELECTED)]
        private IWebElement _recipeMethodSelected;

        [FindsBy(How = How.Id, Using = COEFFICIENT)]
        private IWebElement _recipeCoef;

        [FindsBy(How = How.ClassName, Using = PRICE)]
        private IWebElement _recipePrice;

        // ______________________________________ Méthodes _______________________________________________

        // Général

        public string GetMenuName()
        {
            var menuName = WaitForElementIsVisible(By.XPath(MENU_TITLE));
            return menuName.Text.Substring("Menu : ".Length);
        }
        public bool IsWeekViewVisible()
        {
            if(isElementVisible(By.XPath(MENU_DATES)))
            {
                _menusDates = WaitForElementIsVisible(By.XPath(MENU_DATES));
                return true;
            }
            else
            {
                return false;
            }
        }

        public MenusDayViewPage SwitchToDayView()
        {
            _dayView = WaitForElementIsVisible(By.XPath(DAY_VIEW));
            _dayView.Click();
            WaitForLoad();

            return new MenusDayViewPage(_webDriver, _testContext);
        }

        public void Export(bool newVersionPrint)
        {
            if(isElementVisible(By.Id(EXPORT_DEV)))
            {
                _exportDev = WaitForElementIsVisible(By.Id(EXPORT_DEV));
                _exportDev.Click();
            }
            else
            {
                _exportPatch = WaitForElementIsVisible(By.XPath(EXPORT_PATCH));
                _exportPatch.Click();
            }

            WaitForLoad();

            _saturday = WaitForElementExists(By.Id(SATURDAY));
            _saturday.SetValue(ControlType.CheckBox, true);

            _sunday = WaitForElementExists(By.Id(SUNDAY));
            _sunday.SetValue(ControlType.CheckBox, true);

            WaitPageLoading();
            WaitForLoad();

            _confirmExport = WaitForElementToBeClickable(By.Id(CONFIRM_EXPORT));
            _confirmExport.Click();
            WaitForLoad();

            if (newVersionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
                ClickPrintButton();
            }

            WaitForDownload();
            Close();
        }

        public FileInfo GetExportExcelFile(FileInfo[] taskFiles, string menuName)
        {
            List<FileInfo> correctDownloadFiles = new List<FileInfo>();
            foreach (var file in taskFiles)
            {
                //  Test REGEX
                if (IsExcelFileCorrect(file.Name, menuName))
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
        
        public void ExportMenus()
        {
            ShowExtendedMenu();

            _exportmenus = WaitForElementIsVisible(By.Id(EXPORTMENUS));
            _exportmenus.Click();
            WaitForLoad();

            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-excel']"));
            ClickPrintButton();

            WaitForDownload();
            Close();
        }

        public bool IsExcelFileCorrect(string filePath, string menuName)
        {
            //export-menu- 2139981860-2020-06-15-2020-07-16.xlsx
            string mois = "(?:0[1-9]|1[0-2])";         // mois
            string space = "(\\s)";                    // Espace
            string annee = "\\d{4}";                   // annee YYYY
            string jour = "[0-3]\\d";                  // jour

            //Regex r = new Regex("^export-menu-" + space + menuName + "-" + annee + "-" + mois + "-" + jour + "-" + annee + "-" + mois + "-" + jour + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            Regex r = new Regex("^export-menu-" + menuName + "-" + annee + "-" + mois + "-" + jour + "-" + annee + "-" + mois + "-" + jour + ".xlsx$", RegexOptions.IgnoreCase | RegexOptions.Singleline);

            Match m = r.Match(filePath);

            return m.Success;
        }

        // Add recipe
        public void AddRecipe(string recipeName, DateTime dateTime)
        {
            // A supprimer une fois la page refaite
            _webDriver.Manage().Window.Size = new System.Drawing.Size(768, 500);

            string date = dateTime.ToString("yyyy-MM-dd");

            _addRecipe = WaitForElementExists(By.Id(ADD_RECIPE));
            _addRecipe.Click();
            _addRecipe.SendKeys(recipeName);
            WaitForLoad();

            // A supprimer une fois la page refaite
            _webDriver.Manage().Window.Size = new System.Drawing.Size(1366, 768);
            WaitForLoad();
            var selectRecipe = WaitForElementIsVisible(By.XPath("//p[text()='" + recipeName + "']"));
            selectRecipe.Click();
            WaitForLoad();

            IWebElement activeDay;
            if(isElementVisible(By.XPath(String.Format(ACTIVE_DAY, date))))
            {
                activeDay = WaitForElementIsVisible(By.XPath(String.Format(ACTIVE_DAY, date)));
            }
            else
            {
                var moisSuivant = _webDriver.FindElement(By.XPath(NEXT_MONTH));
                moisSuivant.Click();
                WaitForLoad();

                activeDay = WaitForElementIsVisible(By.XPath(String.Format(ACTIVE_DAY, date)));
            }

            activeDay.Click();
            if(isElementVisible(By.XPath(CONFIRM_ADD)))
            {
                _confirmAdd = WaitForElementIsVisible(By.XPath(CONFIRM_ADD));
            }
            else
            {
                _confirmAdd = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[2]/div/form/div[2]/button[2]"));
            }
            _confirmAdd.Click();
            WaitForLoad();
        }

        // Tableau recipes
        public void ClickOnNextWeek()
        {
            _rightArrow = WaitForElementIsVisible(By.XPath(RIGHT_ARROW));
            _rightArrow.Click();
            WaitForLoad();
        }

        public void ClickOnPreviousWeek()
        {
            _leftArrow = WaitForElementIsVisible(By.XPath(LEFT_ARROW));
            _leftArrow.Click();
            WaitForLoad();
        }

        public bool IsRecipeDisplayed(DateTime dateTime, string recipeName, bool changeView = false)
        {
            try
            {
                if (changeView)
                {
                    _rightArrow = WaitForElementIsVisible(By.XPath(RIGHT_ARROW));
                    _rightArrow.Click();
                    WaitLoading();
                }

                var numJour = GetDayForAddedRecipe(dateTime);

                // Récupération de la cellule dans laquelle la recette a été ajoutée
                _webDriver.FindElement(By.XPath(String.Format(RECIPE_ADDED, numJour+1, recipeName)));

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (changeView)
                {
                    _leftArrow = WaitForElementIsVisible(By.XPath(LEFT_ARROW));
                    _leftArrow.Click();
                    WaitLoading();
                }
            }
        }

        public void ClickOnRecipe(DateTime dateTime, string recipeName)
        {
            var numJour = GetDayForAddedRecipe(dateTime);

            // Récupération de la cellule dans laquelle la recette a été ajoutée
            var recipe = WaitForElementIsVisible(By.XPath(String.Format(RECIPE_ADDED, numJour + 1, recipeName)));
            recipe.Click();
            WaitForLoad();
        }

        public Double GetDayForAddedRecipe(DateTime dateTime)
        {
            // Récupération de la cellule dont la date correspond à la date d'ajout de la recette
            if (!isElementVisible(By.XPath(String.Format(DATE_TABLEAU, dateTime.ToString("dd/MM/yyyy")))))
            {
                throw new Exception("Date non affichée");
            }
            var dateDuJour = WaitForElementExists(By.XPath(String.Format(DATE_TABLEAU, dateTime.ToString("dd/MM/yyyy"))));
            var numJour = dateDuJour.GetAttribute("onclick");

            // Le numéro du jour récupéré correspond à un jour de la semaine
            return Double.Parse(numJour.Substring(numJour.IndexOf("(") + 1, 1));
        }

        public void DeleteRecipe()
        {
            _removeRecipe = WaitForElementIsVisible(By.XPath(REMOVE_BTN));
            _removeRecipe.Click();
            WaitForLoad();
        }

        public void DeleteMultiRecipes()
        {
            _deleteMultiRecipe = WaitForElementIsVisible(By.XPath(MULTI_DELETE));
            _deleteMultiRecipe.Click();
            WaitForLoad();
        }

        public void DuplicateRecipe(DateTime day)
        {
            bool isFound = false;

            _duplicateRecipe = WaitForElementIsVisible(By.XPath(DUPLICATE_BTN));
            _duplicateRecipe.Click();
            WaitForLoad();

            //try
            //{
                var date = WaitForElementIsVisible(By.Id("date-picker-start"));
                date.SetValue(ControlType.DateTime, day);
                date.SendKeys(Keys.Tab);
            //}
            //catch 
            //{
            //    var dayValue = day.Day.ToString();
            //    var otherDays = _webDriver.FindElements(By.XPath(string.Format(DUPLICATE_DATE, dayValue)));

            //    foreach (var otherDay in otherDays)
            //    {
            //        if (!otherDay.GetAttribute("class").Contains("disabled"))
            //        {
            //            otherDay.Click();
            //            isFound = true;
            //            break;
            //        }
            //    }

            //    // Si le jour n'est pas trouvé dans l'affichage courant, on va sur le mois suivant
            //    if (!isFound)
            //    {
            //        _flecheDuplicate = WaitForElementIsVisible(By.XPath(FLECHE_DUPLICATE));
            //        _flecheDuplicate.Click();

            //        var otherDay = _webDriver.FindElement(By.XPath(string.Format(DUPLICATE_DATE, day)));
            //        otherDay.Click();
            //    }

            //}



            _confirmDuplicate = WaitForElementIsVisible(By.XPath(CONFIRM_DUPLICATE));
            _confirmDuplicate.Click();
            WaitForLoad();
        }

        public void DuplicateMultiRecipe(DateTime day)
        {
            bool isFound = false;

            //_duplicateRecipe = WaitForElementIsVisible(By.XPath(DUPLICATE_BTN));//*[@id="topTabWeekly"]/div[1]/a[2]
            _duplicateRecipe = WaitForElementIsVisible(By.XPath("//*[@id=\"topTabWeekly\"]/div[1]/a[2]"));
            _duplicateRecipe.Click();
            WaitForLoad();

            try
            {
                var date = WaitForElementIsVisible(By.Id("date-picker-start"));
                date.SetValue(ControlType.DateTime, day);
                date.SendKeys(Keys.Tab);
            }
            catch
            {
                var dayValue = day.Day.ToString();
                var otherDays = _webDriver.FindElements(By.XPath(string.Format(DUPLICATE_DATE, dayValue)));

                foreach (var otherDay in otherDays)
                {
                    if (!otherDay.GetAttribute("class").Contains("disabled"))
                    {
                        otherDay.Click();
                        isFound = true;
                        break;
                    }
                }

                // Si le jour n'est pas trouvé dans l'affichage courant, on va sur le mois suivant
                if (!isFound)
                {
                    _flecheDuplicate = WaitForElementIsVisible(By.XPath(FLECHE_DUPLICATE));
                    _flecheDuplicate.Click();

                    var otherDay = _webDriver.FindElement(By.XPath(string.Format(DUPLICATE_DATE, day)));
                    otherDay.Click();
                }

            }

            _confirmDuplicate = WaitForElementIsVisible(By.XPath(CONFIRM_DUPLICATE));
            _confirmDuplicate.Click();
            WaitForLoad();
        }

        public void SelectMultiRecipes()
        {
            _selectRecipe = WaitForElementIsVisible(By.XPath(SELECT_BTN));
            _selectRecipe.Click();
            WaitForLoad();
        }

        public RecipeGeneralInformationPage EditRecipe()
        {
            _editRecipe = WaitForElementIsVisible(By.XPath(SEE_RECIPE_DETAIL));
            _editRecipe.Click();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            WaitForLoad();

            return new RecipeGeneralInformationPage(_webDriver, _testContext);
        }

        public void CloseModale()
        {
            _close = WaitForElementIsVisible(By.XPath(CLOSE));
            _close.Click();
            WaitForLoad();
        }

        public string GetRecipeMethodValue()
        {
            //On récupère la valeur de la méthode car pas de selected dans les options  
            _recipeMethodSelected = SolveVisible(METHOD_SELECTED_2);
            var selectedValue = _recipeMethodSelected.GetAttribute("value");
            return selectedValue;
        }

        public void SetRecipeMethod(string method)
        {
            _recipeMethod = SolveVisible(METHOD, true);
            _recipeMethod.SetValue(ControlType.DropDownList, method);
            WaitForLoad();
        }

        public string GetRecipeCoefficient()
        {
            _recipeCoef = SolveVisible(COEFFICIENT, true);
            return _recipeCoef.GetAttribute("value");
        }

        public void SetRecipeCoefficient(string coef)
        {
            _recipeCoef = SolveVisible(COEFFICIENT, true);
            _recipeCoef.SetValue(ControlType.TextBox, coef);

            IsPriceChanged();
            WaitForLoad();
        }

        public string GetRecipePrice()
        {
            _recipePrice = WaitForElementExists(By.ClassName(PRICE));
            return _recipePrice.Text;
        }

        public void IsPriceChanged()
        {
            bool hasChanged = false;
            int cpt = 1;
            string initPrice = GetRecipePrice();

            while(!hasChanged && cpt <= 10)
            {
                if(!initPrice.Equals(GetRecipePrice()))
                {
                    hasChanged = true;
                }
                else
                {
                    cpt++;
                    WaitForLoad();
                }
            }
        }

        public PrintReportPage Print(bool newVersionPrint)
        {
            var _printBtn = WaitForElementIsVisible(By.XPath("//*[@id=\"print\"]")); 
            _printBtn.Click();
            WaitForLoad();

            if (newVersionPrint)
            {
                IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
                ClickPrintButton();
            }

            //Switch the driver to the newly created tab
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }

        public bool CheckUnselectedDaysGris(int daySelected)
         {
            var _listOfDaysWeekView = _webDriver.FindElements(By.XPath(DAYS_LIST_WEEK_VIEW));
            int index = 1;
            //int activeDate = (int)DateUtils.Now.AddDays(1).DayOfWeek;
            foreach(var day in _listOfDaysWeekView)
            {
                //if (index == activeDate) continue;
                var test1 = day.GetAttribute("class");
                if (index != daySelected && !day.GetAttribute("class").Contains("text-center-gray disabled"))
                {
                    return false;
                }
                index++;
            }
            return true;
        }
        public bool CheckUnselectedHaveRightIcon(int daySelected)
        {
            var _listOfelementsDaysWeekView = _webDriver.FindElements(By.XPath(DAYS_ELEMENT_LIST_WEEK_VIEW));
            int index = 1;
            foreach (var day in _listOfelementsDaysWeekView)
            {
                if (index != daySelected && isElementExists(By.XPath("/html/body/div[3]/div[1]/div[3]/div/div[1]/table/tbody/tr[3]/td[" +index.ToString()+"]/p/span")) == false)
                {
                    return false;
                }
            }
            return true;
        }
        public int AddRecipeForManyDatetime(string recipeName, DateTime dateTime)
        {
            // A supprimer une fois la page refaite
            _webDriver.Manage().Window.Size = new System.Drawing.Size(768, 500);

            _addRecipe = WaitForElementExists(By.Id(ADD_RECIPE));
            _addRecipe.Click();
            _addRecipe.SendKeys(recipeName);
            WaitForLoad();
            // A supprimer une fois la page refaite
            _webDriver.Manage().Window.Size = new System.Drawing.Size(1366, 768);
            var selectRecipe = WaitForElementIsVisible(By.XPath("//p[text()='" + recipeName + "']"));
            selectRecipe.Click();
            WaitForLoad();

            // Trouver le début de la semaine (lundi)
            int daysSinceMonday = (int)dateTime.DayOfWeek - 1; // Lundi = 0
            if (daysSinceMonday < 0) daysSinceMonday = 6;   // Dimanche = 6

            DateTime startOfWeek = dateTime.AddDays(-daysSinceMonday);
            DateTime endOfWeek = startOfWeek.AddDays(6); // Fin de semaine = dimanche
            int nb_jdays_selected = 0;
            // Itérer à travers les jours de la semaine et retourner les jours selectionnées
            for (DateTime date = startOfWeek; date <= endOfWeek; date = date.AddDays(1))
            {                
                if (isElementVisible(By.XPath(String.Format(ACTIVE_SELECTED_DAY, date.ToString("yyyy-MM-dd")))))
                {
                    var activeDay = WaitForElementIsVisible(By.XPath(String.Format(ACTIVE_SELECTED_DAY, date.ToString("yyyy-MM-dd"))));
                    activeDay.Click();
                    nb_jdays_selected++;
                }

                //else
                //{
                //    var moisSuivant = _webDriver.FindElement(By.XPath(NEXT_MONTH));
                //    moisSuivant.Click();
                //    WaitForLoad();
                //    activeDay = WaitForElementIsVisible(By.XPath(String.Format(ACTIVE_SELECTED_DAY, date.ToString("yyyy-MM-dd"))));
                //}                
            }
            if (isElementVisible(By.XPath(CONFIRM_ADD)))
            {
                _confirmAdd = WaitForElementIsVisible(By.XPath(CONFIRM_ADD));
            }
            else
            {
                _confirmAdd = WaitForElementIsVisible(By.XPath("//*[@id=\"modal-1\"]/div[2]/div/form/div[2]/button[2]"));
            }
            _confirmAdd.Click();
            WaitForLoad();
            return nb_jdays_selected;
        }
    }
}


