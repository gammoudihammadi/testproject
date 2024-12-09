using DocumentFormat.OpenXml.Office2010.ExcelAc;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newrest.Winrest.FunctionalTests.PageObjects.Shared;
using Newrest.Winrest.FunctionalTests.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.Extensions;
using OpenQA.Selenium.Support.PageObjects;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Policy;
using System.Threading;
using Keys = OpenQA.Selenium.Keys;

namespace Newrest.Winrest.FunctionalTests.PageObjects.ToDoList.Tasks
{
    public class TasksPage : PageBase
    {
        public TasksPage(IWebDriver webDriver, TestContext testContext) : base(webDriver, testContext)
        {
        }

        //---------------------------------------CONSTANTES---------------------------------------------------------------

        public const string PLUS_BTN = "//div[@class=\"dropdown dropdown-add-button\"]/button";
        public const string CREATE_NEW_TASK = "//a[text()=\"New task\"]";
        public const string NUMBER_TASKS_IN_HEADER = "/html/body/div[2]/div/div[2]/div/div[1]/h1/span";
        public const string NAME = "Name";
        public const string SITES = "SelectedSites_ms";
        public const string CUSTOMER = "CustomerId";
        public const string ENFORCE_ORDER = "EnforceOrder";
        public const string CREATE_BTN = "/html/body/div[3]/div/div/div[2]/div/div/form/div/div/div[2]/div/button[2]";
        public const string BACK_TO_LIST = "/html/body/div[2]/a/span[2]";
        public const string DELETE_ICON = "/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[4]/a";
        public const string DELETE_CONFIRM = "//a[@id='dataConfirmOK'][text()='Delete']";
        public const string CREATE_FROM_OTHER_TASK_BUTTON = "//a[text()=\"Create From Other Task\"]";
        public const string FIRST_NAME_TASK_GRID = "/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[3]/table/tbody/tr/td[2]";
        public const string FIRST_SITE_GRID = "/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[3]/table/tbody/tr/td[3]/span";
        public const string FIRST_CUSTOMER_GRID = "/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div/div/div[3]/table/tbody/tr[*]/td[4]";
        public const string DESCRIPTION = "Description";
        public const string LIST_SITES_SELECTED = "/html/body/div[3]/div/div/div/div/div[2]/div/div/form/div/div/div[1]/div[1]/div[2]/div/div/div/div[2]/select/option[@selected]";
        public const string VEHICULE_ID = "VehiculeId";
        public const string PRINT_QR_CODE = "/html/body/div[3]/div/div/div/div/div[1]/a";
        public const string PRINT_BTN = "printTaskButton";
        public const string LIST_QR_CODE_SIZE = "/html/body/div[4]/div/div/div[2]/div/form/div[1]/select";
        public const string LIST_QR_CODE_TYPE = "Types";
        public const string CUSTOMER_SELECTED = "/html/body/div[3]/div/div/div/div/div[2]/div/div/form/div/div/div[1]/div[1]/div[3]/div/select/option[@selected]";
        public const string FOLD_UNFOLD = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[1]/span";
        public const string STEP_FOLD_LINE = "/html/body/div[2]/div/div[2]/div/div[2]/div[2]/div[2]/div/table/tbody[2]/tr";
        //---------------------------------------CREATE FROM OTHER TASK CONSTANTES---------------------------------------------------------------
        public const string TASKS = "Tasks";
        public const string NEW_TASKS_NAME = "taskName";
        public const string DUPLICATE_SCHEDULE = "DuplicateScheduler";
        public const string DUPLICATE_TASK_BUTTON = "DuplicateTaskButton";
        //---------------------------------------FILTERS CONSTANTES-------------------------------------------------------
        public const string SEARCH = "SearchString";
        public const string SHOW_ALL = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[3]/input[1]";
        public const string SHOW_ONLY_ACTIVE = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[3]/input[2]";
        public const string SHOW_ONLY_INACTIVE = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[3]/input[3]";
        public const string SITE_FILTER = "FilterSites_ms";
        public const string CUSTOMER_FILTER = "FilterCustomers_ms";
        public const string RESET_FILTER = "/html/body/div[2]/div/div[1]/div[1]/div/form/div[1]/a";
        public const string LIST_SITE_FILTER = "/html/body/div[10]/ul/li[*]/label/input";
        public const string LIST_CUSTOMER_FILTER = "/html/body/div[11]/ul/li[*]/label/input";
        public const string SCAN_ONLY = "ScanOnly";
        public const string FIRST_TASK_NAME = "//*[@id=\"list-item-with-action\"]/div[2]/div[1]/div/div[3]/table/tbody/tr/td[2]";
        public const string IS_ACTIVE = "IsActive";

        // ____________________________________ Variables _______________________________________________
        [FindsBy(How = How.XPath, Using = PLUS_BTN)]
        private IWebElement _plusBtn;

        [FindsBy(How = How.XPath, Using = CREATE_NEW_TASK)]
        private IWebElement _createBtn;

        [FindsBy(How = How.Id, Using = NAME)]
        private IWebElement _name;

        [FindsBy(How = How.Id, Using = SITES)]
        private IWebElement _sites;

        [FindsBy(How = How.Id, Using = CUSTOMER)]
        private IWebElement _customer;

        [FindsBy(How = How.XPath, Using = SHOW_ALL)]
        private IWebElement _showAll;

        [FindsBy(How = How.XPath, Using = SHOW_ONLY_ACTIVE)]
        private IWebElement _showOnlyActive;

        [FindsBy(How = How.XPath, Using = SHOW_ONLY_INACTIVE)]
        private IWebElement _showOnlyInActive;

        [FindsBy(How = How.Id, Using = SEARCH)]
        private IWebElement _search;

        [FindsBy(How = How.XPath, Using = CREATE_FROM_OTHER_TASK_BUTTON)]
        private IWebElement _createFromOtherTasksBtn;

        [FindsBy(How = How.Id, Using = TASKS)]
        private IWebElement _tasks;

        [FindsBy(How = How.Id, Using = NEW_TASKS_NAME)]
        private IWebElement _newTaskName;

        [FindsBy(How = How.Id, Using = DUPLICATE_SCHEDULE)]
        private IWebElement _duplicateSchedule;

        [FindsBy(How = How.Id, Using = DUPLICATE_TASK_BUTTON)]
        private IWebElement _duplicateTaskButton;

        [FindsBy(How = How.Id, Using = ENFORCE_ORDER)]
        private IWebElement _EnforceOrder;

        [FindsBy(How = How.Id, Using = DESCRIPTION)]
        private IWebElement _description;

        [FindsBy(How = How.XPath, Using = PRINT_QR_CODE)]
        private IWebElement _printQRCode;

        [FindsBy(How = How.Id, Using = PRINT_BTN)]
        private IWebElement _printBtn;

        [FindsBy(How = How.XPath, Using = LIST_QR_CODE_SIZE)]
        private IWebElement _qrCodeSize;

        [FindsBy(How = How.XPath, Using = FOLD_UNFOLD)]
        private IWebElement _foldUnfold;

        [FindsBy(How = How.XPath, Using = STEP_FOLD_LINE)]
        private IWebElement _stepUnderfoldUnfoldLine;

        [FindsBy(How = How.XPath, Using = SCAN_ONLY)]
        private IWebElement _scanOnly;
        [FindsBy(How = How.Id, Using = IS_ACTIVE)]
        private IWebElement _isActive;
        public enum FilterType
        {
            Search,
            ShowAll,
            ShowOnlyActive,
            ShowOnlyInactive,
            Sites,
            Customers
        }

        public void Filter(FilterType filterType, object value)

        {
            switch (filterType)
            {
                case FilterType.Search:

                    _search = WaitForElementIsVisibleNew(By.Id(SEARCH));
                    _search.Clear();
                    _search.SetValue(ControlType.TextBox, value);
                    break;

                case FilterType.ShowAll:
                    _showAll = WaitForElementIsVisibleNew(By.XPath(SHOW_ALL));
                    _showAll.SetValue(ControlType.CheckBox, value);
                    break;

                case FilterType.ShowOnlyActive:
                    _showOnlyActive = WaitForElementIsVisibleNew(By.XPath(SHOW_ONLY_ACTIVE));
                    _showOnlyActive.SetValue(ControlType.CheckBox, value);
                    break;

                case FilterType.ShowOnlyInactive:
                    _showOnlyInActive = WaitForElementIsVisibleNew(By.XPath(SHOW_ONLY_INACTIVE));
                    _showOnlyInActive.SetValue(ControlType.CheckBox, value);
                    break;

                case FilterType.Sites:
                    ComboBoxSelectById(new ComboBoxOptions(SITE_FILTER, (string)value));
                    break;

                case FilterType.Customers:
                    ComboBoxSelectById(new ComboBoxOptions(CUSTOMER_FILTER, (string)value));
                    break;

                default:
                    break;

            }
            WaitPageLoading();
            Thread.Sleep(3000);
        }
        public object GetFilterValue(FilterType filterType)
        {
            switch (filterType)
            {
                case FilterType.Search:
                    _search = WaitForElementIsVisible(By.Id(SEARCH));
                    return _search.GetAttribute("value");

                case FilterType.ShowAll:
                    _showAll = WaitForElementIsVisible(By.XPath(SHOW_ALL));
                    return _showAll.Selected;

                case FilterType.ShowOnlyActive:
                    _showOnlyActive = WaitForElementIsVisible(By.XPath(SHOW_ONLY_ACTIVE));
                    return _showOnlyActive.Selected;

                case FilterType.ShowOnlyInactive:
                    _showOnlyInActive = WaitForElementIsVisible(By.XPath(SHOW_ONLY_INACTIVE));
                    return _showOnlyInActive.Selected;
            }
            return null;
        }
        public void ResetFilters()
        {
            var resetFilter = WaitForElementIsVisibleNew(By.XPath(RESET_FILTER));
            resetFilter.Click();

            WaitForLoad();
            // batch passage à minuit-2h UTC Renaud
            if (DateUtils.IsBeforeMidnight())
            {
                //pas de date
            }
        }
        public void ShowPlusScheduler()
        {
            _plusBtn = WaitForElementIsVisibleNew(By.XPath(PLUS_BTN));
            _plusBtn.Click();
            WaitForLoad();
        }
        public TasksGeneralInfo CreateNewTask(string nameInput, string sitesInput, string customerInput, bool allsites = false, bool isActive = true, bool isScanOnly = false)
        {
            ShowPlusScheduler();
            _createBtn = WaitForElementIsVisible(By.XPath(CREATE_NEW_TASK));
            _createBtn.Click();
            WaitForLoad();

            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, nameInput);
            WaitForLoad();
            if (allsites)
            {
                var site = WaitForElementExists(By.Id(SITES));
                site.Click();

                var selectAll = WaitForElementIsVisible(By.XPath("/html/body/div[13]/div/ul/li[1]/a/span[2]"));
                selectAll.Click();
            }
            else
            {
                ComboBoxSelectById(new ComboBoxOptions(SITES, sitesInput, false));
            }


            _customer = WaitForElementIsVisible(By.Id(CUSTOMER));
            _customer.SendKeys(customerInput);
            WaitForLoad();

            if (isScanOnly)
            {
                _scanOnly = WaitForElementExists(By.Id(SCAN_ONLY));
                _scanOnly.Click();
                WaitForLoad();
            }

            if (isActive == false)
            {
                _isActive = WaitForElementExists(By.Id(IS_ACTIVE));
                _isActive.Click();
                WaitForLoad();
            }

            var createBtn = WaitForElementIsVisible(By.XPath(CREATE_BTN));
            createBtn.Click();

            WaitPageLoading();
            return new TasksGeneralInfo(_webDriver, _testContext);
        }
        public TasksGeneralInfo CreateNewTaskWithMultipleSite(string nameInput, List<string> sitesInput, string customerInput)
        {
            ShowPlusScheduler();

            _createBtn = WaitForElementIsVisible(By.XPath(CREATE_NEW_TASK));
            _createBtn.Click();
            WaitForLoad();

            _name = WaitForElementIsVisible(By.Id(NAME));
            _name.SetValue(ControlType.TextBox, nameInput);
            WaitForLoad();

            foreach (var siteInput in sitesInput)
            {
                ComboBoxSelectMultipleById(new ComboBoxOptions(SITES, siteInput, false));
            }

            _customer = WaitForElementIsVisible(By.Id(CUSTOMER));
            _customer.SendKeys(customerInput);
            WaitForLoad();

            var createBtn = WaitForElementIsVisible(By.XPath(CREATE_BTN));
            createBtn.Click();

            WaitPageLoading();
            return new TasksGeneralInfo(_webDriver, _testContext);
        }
        public void ComboBoxSelectMultipleById(ComboBoxOptions cbOpt)
        {
            IWebElement input = WaitForElementIsVisible(By.Id(cbOpt.XpathId));
            input.Click();
            WaitForLoad();

            //if (cbOpt.ClickCheckAllAtStart)
            //{
            //    var checkAllVisible = SolveVisible("//*/span[text()='Check all']");
            //    Assert.IsNotNull(checkAllVisible);
            //    checkAllVisible.Click();
            //}
            //else if (cbOpt.ClickUncheckAllAtStart)
            //{
            //    var uncheckAllVisible = SolveVisible("//*/span[text()='Uncheck all']");
            //    Assert.IsNotNull(uncheckAllVisible);
            //    uncheckAllVisible.Click();
            //}

            if (cbOpt.IsUsedInFilter)
            {
                WaitPageLoading();
                WaitForLoad();
            }
            else if (cbOpt.ClickUncheckAllAtStart)
            {
                WaitForLoad();
            }

            bool selectionWasModified = false;

            if (cbOpt.SelectionValue != null)
            {
                var searchVisible = SolveVisible("//*/input[@type='search']");
                Assert.IsNotNull(searchVisible);
                if (cbOpt.IsUsedInFilter)
                {
                    searchVisible.SetValue(ControlType.TextBox, cbOpt.SelectionValue);
                }
                else
                {
                    _webDriver.ExecuteJavaScript("arguments[0].value = arguments[1]; arguments[0].dispatchEvent(new Event('input'));", searchVisible, cbOpt.SelectionValue);
                }
                // on ne clique pas de checkbox donc pas de rechargement de page ici
                Thread.Sleep(3000);
                WaitForLoad();

                var select = SolveVisible("//*/label[contains(@for, 'ui-multiselect')]/span[contains(text(),'" + cbOpt.SelectionValue + "')]");
                Assert.IsNotNull(select, "Pas de sélection de " + cbOpt.SelectionValue);

                if (cbOpt.ClickCheckAllAfterSelection)
                {
                    var checkAllVisible = SolveVisible("//*/span[text()='Check all']");
                    Assert.IsNotNull(checkAllVisible);
                    checkAllVisible.Click();
                }
                else if (cbOpt.ClickUncheckAllAfterSelection)
                {
                    var uncheckAllVisible = SolveVisible("//*/span[text()='Uncheck all']");
                    Assert.IsNotNull(uncheckAllVisible);
                    uncheckAllVisible.Click();
                }
                else
                {
                    select.Click();
                }
                selectionWasModified = true;
            }

            if (selectionWasModified)
            {
                if (cbOpt.IsUsedInFilter)
                {
                    WaitPageLoading();
                    WaitForLoad();
                }
                else
                {
                    WaitForLoad();
                }
            }

            input = WaitForElementIsVisible(By.Id(cbOpt.XpathId));

            try
            {
                input.SendKeys(Keys.Enter);
            }
            catch
            {
                //Silent catch: sometimes there's no associated action with "enter" key
                input.Click();
            }

            WaitForLoad();
        }
        public string GetNumberOfTasksInHeader()
        {
            var numberTasks = WaitForElementIsVisibleNew(By.XPath(NUMBER_TASKS_IN_HEADER));
            return numberTasks.Text;
        }
        public void BackToList()
        {
            var backToListBtn = WaitForElementIsVisibleNew(By.XPath(BACK_TO_LIST));
            backToListBtn.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        public void DeleteTask()
        {
            if (isElementExists(By.XPath(DELETE_ICON)))
            {
                WaitForLoad();
                var delete = _webDriver.FindElement(By.XPath(DELETE_ICON));
                Actions actions = new Actions(_webDriver);
                actions.MoveToElement(delete).Perform();
                delete.Click();
            }
            WaitPageLoading();
            WaitForLoad();

            var deleteConfirm = WaitForElementIsVisible(By.XPath(DELETE_CONFIRM));
            deleteConfirm.Click();
            WaitPageLoading();
            WaitForLoad();
        }
        
        
        public void CreateFromOtherTasks(string task, string newTaskName, bool duplicateSchedule)
        {
            ShowPlusScheduler();
            WaitForLoad();

            _createFromOtherTasksBtn = WaitForElementIsVisibleNew(By.XPath(CREATE_FROM_OTHER_TASK_BUTTON));
            _createFromOtherTasksBtn.Click();
            WaitForLoad();

            _tasks = WaitForElementIsVisibleNew(By.Id(TASKS));
            
            _tasks.Click();
            var list = _webDriver.FindElements(By.XPath("//*[@id=\"Tasks\"]/option"));
            var task_to_choise = WaitForElementIsVisibleNew(By.XPath(string.Format("//*[@id=\"Tasks\"]/option[{0}]", list.Count)));
            task_to_choise.Click();


            _newTaskName = WaitForElementIsVisibleNew(By.Id(NEW_TASKS_NAME));
            _newTaskName.SetValue(ControlType.TextBox, newTaskName);

            _duplicateSchedule = WaitForElementIsVisibleNew(By.Id(DUPLICATE_SCHEDULE));
            _duplicateSchedule.SetValue(ControlType.CheckBox, duplicateSchedule);

            _duplicateTaskButton = WaitForElementIsVisibleNew(By.Id(DUPLICATE_TASK_BUTTON));
            _duplicateTaskButton.Click();
            WaitForLoad();
        }
        public bool VerifySearchFilter(string name)
        {
            WaitPageLoading();
            var firstNameTaskIInGrid = WaitForElementIsVisible(By.XPath(FIRST_NAME_TASK_GRID));
            if (!firstNameTaskIInGrid.Text.Contains(name))
            {
                return false;
            }
            return true;
        }
        public bool VerifyCustomerFilter(string customer)
        {
            var firstCustomerInGrid = _webDriver.FindElements(By.XPath(FIRST_CUSTOMER_GRID));
            foreach (var firstCustomer in firstCustomerInGrid)
            {
                if (firstCustomer.Text != customer)
                {
                    return false;
                }
            }
            
            return true;
        }
        public bool VerifySiteFilter(string site)
        {
            var firstSiteInGrid = WaitForElementIsVisible(By.XPath(FIRST_SITE_GRID));
            var title = firstSiteInGrid.GetAttribute("title");
            if (!title.Contains(site))
            {
                return false;
            }
            return true;
        }
        public TasksGeneralInfo SelectFirstTask()
        {
            WaitForLoad();
            WaitForElementIsVisible(By.XPath(FIRST_NAME_TASK_GRID)).Click();
            
            WaitPageLoading();
            return new TasksGeneralInfo(_webDriver, _testContext);
        }
        
        public int GetNumberSelectedSiteFilter()
        {
            var listSitesSelectedFilters = _webDriver.FindElements(By.XPath(LIST_SITE_FILTER));
            var numberSitesSelectedSite = listSitesSelectedFilters
               .Where(p => p.Selected == true).Count();

            return numberSitesSelectedSite;
        }
        public int GetNumberSelectedCustomerFilter()
        {
            var listCustiomersSelectedFilters = _webDriver.FindElements(By.XPath(LIST_CUSTOMER_FILTER));
            var numberCustomersSelectedSite = listCustiomersSelectedFilters
               .Where(p => p.Selected == true).Count();

            return numberCustomersSelectedSite;
        }
        public void PrintQRCode()
        {
            _printQRCode = WaitForElementIsVisible(By.XPath(PRINT_QR_CODE));
            _printQRCode.Click();
            WaitLoading();

            _qrCodeSize = WaitForElementIsVisible(By.XPath(LIST_QR_CODE_SIZE));
            _qrCodeSize.SetValue(ControlType.TextBox, "Small");
            _qrCodeSize.SendKeys(Keys.Tab);
            WaitLoading();

            var _qrCodeType = WaitForElementIsVisible(By.Id(LIST_QR_CODE_TYPE));
            _qrCodeType.SetValue(ControlType.TextBox, "Scan on Tablet");
            _qrCodeType.SendKeys(Keys.Tab);
            WaitLoading();

            _printBtn = WaitForElementIsVisible(By.Id(PRINT_BTN));
            _printBtn.Click();
            WaitLoading();
        }
        public PrintReportPage PrintResults()
        {
            IsFileLoaded(By.CssSelector("[target='_blank'][class='far fa-file-pdf']"));
            ClickPrintButton();
            WaitPageLoading();
            WaitForLoad();

            //Results are opened in a new tab, switch the driver to the newly created one
            var wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(60));
            wait.Until((driver) => driver.WindowHandles.Count > 1);
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles[1]);

            return new PrintReportPage(_webDriver, _testContext);
        }
        public PrintReportPage PrintItemGenerique(TasksPage tasksPage)
        {
            // Lancement du Print
            PrintReportPage reportPage;
            reportPage = tasksPage.PrintResults();
            // le zip se télécharge directement, alors que le pdf s'ouvre dans un nouveau onglet
            var isReportGenerated = reportPage.IsReportGenerated();
            reportPage.Close();
            tasksPage.ClickPrintButton();

            //Assert
            Assert.IsTrue(isReportGenerated, "L'application n'a pas pu générer le fichier attendu.");
            return reportPage;
        }
        public void FoldAndUnfold()
        {
            _foldUnfold = WaitForElementIsVisible(By.XPath(FOLD_UNFOLD));
            _foldUnfold.Click();
            Thread.Sleep(1500);
        }
        public bool VerifyFoldAndUnfold()
        {
            WaitForLoad();
            if(!isElementVisible(By.XPath(STEP_FOLD_LINE)))
            {
                return false;
            }
            return true;
        }
        public string GetFirstTaskName()
        {
            var firstNameTaskIInGrid = WaitForElementIsVisible(By.XPath(FIRST_NAME_TASK_GRID));
            return firstNameTaskIInGrid.Text;
           
        }
    }
}
